# Outlook MCP Server

A **Model Context Protocol** server bridging Claude to **Microsoft Outlook + Teams** (email, calendar, contacts, tasks, files, Teams meeting recordings + transcripts), deployed on **Cloudflare Workers**.

Fork this repo, deploy to your own Cloudflare account, register a Microsoft Azure AD app, point Claude.ai at your worker, and Claude can read and write your Microsoft 365 data through natural language.

Built on [`@bashco/mcp-toolkit`](https://github.com/doublebash/mcp-toolkit) — OAuth, per-client bearer tokens, rate limiting, structured logging, typed tool dispatch are all handled by the shared library.

## What Claude gets — 38 tools across 7 domains

- **Mail**: list emails, read email, search, reply, forward, delete, send, move between folders, create draft, update draft, send draft, schedule send
- **Calendar**: list events, list event occurrences, create, update, delete, cancel event, respond to event
- **Contacts**: list, create contact
- **Tasks**: list task lists, list tasks, create task
- **Files**: list files, share file
- **Teams meetings**: list recent recordings (the discovery starting point — finds meetings that have content in the past N days, no inputs needed), find online meeting, list meeting recordings, list meeting transcripts, get transcript content. Each per-meeting tool accepts any of `meeting_id`, `calendar_event_id`, or `join_url` — so scheduled meetings (resolved via event), ad-hoc / Meet-now calls (resolved via join URL pasted from the Teams chat), and direct id lookups all work.
- **Conversation**: get conversation (full thread)
- **Settings**: get mailbox settings, set out-of-office

Full live catalogue at the `tools/list` MCP endpoint after deploy.

## How auth works

Two layers:

1. **Claude.ai ↔ your worker** — standard MCP OAuth 2.0 + PKCE flow. Each Claude client gets a unique bearer token; your `MCP_APPROVAL_CODE` is what *you* paste at `/authorize` once to mint that bearer.
2. **Your worker ↔ Microsoft Graph** — proxied OAuth. You authorise Microsoft *once* by visiting `/oauth/start` on your deployed worker; refresh tokens are stored encrypted at rest in Cloudflare KV. Refresh happens automatically.

## Setup — deploy your own copy

### Prerequisites

- A [Cloudflare account](https://dash.cloudflare.com/sign-up) (free plan works)
- [Wrangler CLI](https://developers.cloudflare.com/workers/wrangler/install-and-update/) installed and logged in (`wrangler login`)
- Node.js 22+
- A Microsoft account (personal, work, or school) with access to register Azure AD apps at [entra.microsoft.com](https://entra.microsoft.com)

### 1. Fork and clone

```bash
git clone https://github.com/<your-username>/outlook-mcp
cd outlook-mcp
npm install
```

### 2. Create the KV namespace

```bash
wrangler kv:namespace create OAUTH_KV
```

Wrangler prints something like:

```
🌀 Creating namespace with title "outlook-mcp-OAUTH_KV"
✨ Success! Add the following to your configuration file:
[[kv_namespaces]]
binding = "OAUTH_KV"
id = "abc123def456..."
```

Edit `wrangler.jsonc` and replace the existing `id` under `kv_namespaces` with what wrangler just printed.

> `wrangler.jsonc` is committed by design — Wrangler and the CI deploy both need
> it, and it holds no secrets (only your KV namespace id, public Azure client id,
> and worker URL). `wrangler.jsonc.example` carries the same structure with
> placeholders if you'd rather start from a clean copy. Real secrets go through
> `wrangler secret put` and never appear in this file.

### 3. Register a Microsoft Azure AD app

1. Go to [entra.microsoft.com → Identity → Applications → App registrations → New registration](https://entra.microsoft.com)
2. **Name**: anything (e.g. "Claude Outlook MCP")
3. **Supported account types**:
   - "Accounts in this organizational directory only" if you want to restrict to one tenant
   - "Accounts in any organizational directory and personal Microsoft accounts" for the broadest support
4. **Redirect URI**: leave blank for now — you'll come back after Step 6
5. Click Register
6. From the app's Overview page, record:
   - **Application (client) ID** → this is your `MICROSOFT_CLIENT_ID`
   - **Directory (tenant) ID** → this is your `MICROSOFT_TENANT_ID` (or use the string `common` for multi-tenant + personal account support)
7. **API permissions** → Add the following Microsoft Graph delegated permissions:
   - `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.ReadWrite`
   - `Contacts.ReadWrite`
   - `Tasks.ReadWrite`
   - `Files.Read.All` (or `Files.ReadWrite.All` if you want write-capable file tools)
   - `User.Read`
   - `offline_access` (required for refresh tokens)
   - `MailboxSettings.ReadWrite`
   - `Sites.Read.All`
   - `OnlineMeetings.Read`
   - `OnlineMeetingRecording.Read.All` — **admin consent required**
   - `OnlineMeetingTranscript.Read.All` — **admin consent required**

   After adding the two `.Read.All` permissions, click **"Grant admin consent for [tenant name]"** on the API permissions page. Without admin consent, the meeting recording / transcript tools will 403.
8. **Certificates & secrets** → New client secret → record the value (in 1Password). This is your `MICROSOFT_CLIENT_SECRET`. You can only see it once — copy immediately.

### 4. Update wrangler.jsonc

Edit `wrangler.jsonc` and replace both:
- `vars.MICROSOFT_CLIENT_ID` — with the Application ID from Step 3.6
- `vars.MICROSOFT_TENANT_ID` — with the Directory ID from Step 3.6 (or `common`)

### 5. Set secrets

Generate a fresh approval code:

```bash
openssl rand -base64 32
```

Store it in a password manager, then push to Cloudflare:

```bash
wrangler secret put MCP_APPROVAL_CODE          # paste the value from above
wrangler secret put MICROSOFT_CLIENT_SECRET    # from Step 3.8
```

| Secret                     | Purpose |
|---|---|
| `MCP_APPROVAL_CODE`        | One-time code you paste at `/authorize` to mint a Claude bearer. Also used as the encryption secret for upstream Microsoft tokens at rest — rotating it invalidates stored tokens and forces clean Microsoft re-auth. |
| `MICROSOFT_CLIENT_SECRET`  | Your Azure AD app's client secret. |
| `SIGNATURE_HTML`           | Optional. Email signature block appended server-side — see [Email signature](#email-signature). |
| `SIGNATURE_LOGO_URL`       | Optional. Publicly reachable HTTPS URL of the signature logo. |

### 6. First deploy (to learn worker URL)

```bash
npm run deploy
```

Wrangler prints your worker URL — something like `https://outlook-mcp.<your-account>.workers.dev`. Save it.

### 7. Update `WORKER_URL` and the Microsoft redirect URI

Two updates needed:

**a) Edit `wrangler.jsonc`** — under `vars`, replace `WORKER_URL` with the URL from Step 6.

**b) In the Azure AD app** (entra.microsoft.com → your app → Authentication → Add a platform → Web), set the redirect URI to `<your-worker-url>/oauth/callback`. Without this, Microsoft will reject the OAuth flow.

Then redeploy:

```bash
npm run deploy
```

### 8. Connect Microsoft (one time)

In your browser, visit `<your-worker-url>/oauth/start`. Paste your `MCP_APPROVAL_CODE`. You'll be redirected to Microsoft to sign in and grant the scopes from Step 3.7. After consenting, your encrypted upstream tokens land in `OAUTH_KV`. Refresh happens automatically thereafter.

You can confirm the connection by visiting `<your-worker-url>/oauth/status` — should say `connected: true`.

### 9. Connect Claude.ai

1. In Claude.ai, go to **Settings → Integrations → Add MCP server**
2. Server URL: `<your-worker-url>/mcp`
3. Claude.ai redirects you to your worker's `/authorize` page
4. Paste your `MCP_APPROVAL_CODE` and confirm
5. You're connected — Claude now has the 38 Outlook + Teams tools available

## Email signature

Optional. When configured, the Worker appends your signature at send time so
the calling agent never has to reproduce it — it can't be paraphrased,
truncated, or forgotten.

Pass `include_signature: true` to any of `send_email`, `schedule_send`,
`reply_to_email`, `forward_email`, `create_draft`, `update_draft`,
`create_reply_draft`, `create_reply_all_draft`, or `create_forward_draft`. It
defaults to `false`, so existing callers are unaffected.

For drafts, the signature is injected at **draft-creation time**, not at send
time — `send_draft` only takes an id and never touches the body. This means the
signed body is what you review before sending. On `update_draft` the flag only
applies when you also pass a new `body` (otherwise there is nothing to sign, and
it would replace the draft with a signature-only body); that case is reported in
the response `notes` rather than silently wiping the draft.

### Calendar invites

`create_calendar_event` and `update_calendar_event` accept the same
`include_signature` flag, appending the signature to the event **description**.
It reuses the same `SIGNATURE_HTML` block as email — including its marketing
call-to-action buttons — so it suits a client-facing invite more than an
internal meeting; that is why it is opt-in per event. `update_calendar_event`
follows the same guard as `update_draft`: the flag only applies when you also
pass a new `description`.

### Setup

```bash
cp signature-block.example.html signature-block.html   # then edit it
wrangler secret put SIGNATURE_HTML < signature-block.html
wrangler secret put SIGNATURE_LOGO_URL                 # paste your HTTPS logo URL
```

`signature-block.html` is **gitignored on purpose**. The signature is
deployment config, not source: a fork that inherited a committed signature
would send mail carrying someone else's name, phone number, and booking links.
Only the placeholder `signature-block.example.html` is committed.

The `__LOGO_URL__` token inside `SIGNATURE_HTML` is substituted with
`SIGNATURE_LOGO_URL` at runtime.

**The logo must be publicly reachable over HTTPS.** Mail clients fetch it from
the *recipient's* machine — it has no access to your network, your Worker's
bindings, or any credential you hold. A private, authenticated, or
`localhost` URL renders as a broken image for everyone. If
`SIGNATURE_LOGO_URL` is unset the `<img>` is dropped entirely rather than
emitting a broken `src`.

If `SIGNATURE_HTML` is unset the flag is a silent no-op — the mail sends
unsigned. An unconfigured deployment never errors.

### Behaviour

- **Forces HTML.** A signature inside a plain-text body renders as visible raw
  markup, so `body_type` is overridden to `html` whenever the flag is set. When
  you explicitly passed `body_type: "text"` the override is reported back in the
  tool response's `notes`, never applied silently.
- **Plain-text bodies are escaped, then newlines become `<br>`**, so your line
  breaks survive the forced HTML switch and stray `<` characters can't become
  markup.
- **Idempotent.** If the body already carries the signature — recognised by the
  Worker's own marker or by the signature's distinctive text — it is not
  appended twice.
- **Replies and forwards** place the signature above the quoted original, not at
  the bottom of the whole thread.
- **Empty body** sends the signature alone, with no leading blank lines.

### Security note

The signature is concatenated **after** the caller's body has been sanitised.
This is deliberate and load-bearing: `sanitizeOutboundHtml` strips every
`style=` attribute (an attribute-only XSS sink), and the signature is built
entirely from inline styles, so routing it through the sanitiser would strip
the logo sizing, the divider, and the CTA buttons.

The two strings have different trust levels. The body is agent-supplied and
untrusted, so it is still fully sanitised. The signature is operator-supplied
deployment config set via `wrangler secret put` — anyone who can set that
secret can already change the Worker outright. See `src/signature.ts`.

## Local development

```bash
cp .dev.vars.example .dev.vars   # fill in MCP_APPROVAL_CODE + MICROSOFT_CLIENT_SECRET; .dev.vars is gitignored
npm test                          # 171 tests via vitest with workers pool
npm run typecheck                 # tsc --noEmit
npm run dev                       # wrangler dev — local at http://localhost:8787
```

## Endpoints

- `GET /.well-known/oauth-authorization-server` — OAuth metadata (public)
- `GET /.well-known/oauth-protected-resource` — Resource metadata (public)
- `GET /authorize` — Approval-code paste page (public)
- `POST /approve` — Approval-code submission (rate-limited)
- `POST /token` — OAuth token exchange (rate-limited)
- `POST /register` — Dynamic client registration per RFC 7591 (rate-limited)
- `GET /oauth/start` — Begin Microsoft OAuth flow (gated by MCP_APPROVAL_CODE)
- `GET /oauth/callback` — Microsoft OAuth redirect target
- `GET /oauth/status` — Check connection state (gated by MCP_APPROVAL_CODE)
- `POST /mcp` — JSON-RPC tool dispatch (bearer-protected, rate-limited)

## Stack

- Cloudflare Workers (compatibility_date `2025-04-28`, `nodejs_compat`)
- TypeScript (strict)
- [Hono](https://hono.dev) v4
- [Zod](https://zod.dev) v4
- Vitest with `@cloudflare/vitest-pool-workers` (171 tests)
- [`@bashco/mcp-toolkit`](https://github.com/doublebash/mcp-toolkit) — shared OAuth/crypto/rate-limit/dispatch plumbing

## Security architecture highlights

- **Two-pass HTML sanitiser** with entity normalisation on outbound email previews
- **SSRF guard** with 32-bit-IP normalisation on outbound HTTP
- **Microsoft Graph odata error envelope parsing** for structured error returns
- **Per-domain tool files** under `src/tools/` (mail, calendar, contacts, tasks, files, meetings, settings) for auditability

## Continuous deployment

`.github/workflows/deploy.yml` runs `vitest run` on every push to `main`, then deploys to Cloudflare. To enable on your fork, set two repository secrets:

- `CLOUDFLARE_API_TOKEN` — create at [dash.cloudflare.com/profile/api-tokens](https://dash.cloudflare.com/profile/api-tokens) (use the "Edit Cloudflare Workers" template)
- `CLOUDFLARE_ACCOUNT_ID` — find at the bottom-right of your Cloudflare dashboard

## Contributing

Issues and PRs welcome at [github.com/doublebash/outlook-mcp](https://github.com/doublebash/outlook-mcp).

For changes to the underlying OAuth/crypto/rate-limit code, the toolkit lives at [github.com/doublebash/mcp-toolkit](https://github.com/doublebash/mcp-toolkit) — file issues there.

## Security

Found a vulnerability? Please **don't** open a public issue. Open a [private security advisory](https://github.com/doublebash/outlook-mcp/security/advisories/new) on GitHub.

## License

[MIT](./LICENSE) — Copyright (c) 2026 Bashar Basheer.
