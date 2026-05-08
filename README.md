# Outlook MCP Server

**Status:** Live  
**Built:** April 2026  
**Worker URL:** https://outlook-mcp.bashar-basheer.workers.dev  
**Project folder:** `/Users/basheerco/Documents/AI Agents & Workflows/Outlook-MCP/outlook-mcp`

---

## What this is

A custom MCP (Model Context Protocol) server that connects Claude directly to your Microsoft Outlook account. It gives Claude full read and write access to your email, calendar, contacts, tasks, mailbox settings, and OneDrive — all controlled through natural language in Claude Desktop.

This is not a plugin or third-party integration. It is infrastructure you own and control, hosted on your own Cloudflare account, authenticated against your own Microsoft Azure app registration.

---

## What Claude can do with it

### Email
- List recent emails from any folder (inbox, sent, drafts, subfolders)
- Read the full content of any email
- Send new emails with CC and BCC
- Reply to emails (reply or reply all)
- Delete emails
- Move emails between folders
- Search across all mail

### Calendar
- List upcoming events within any date range
- Create events with title, time, timezone, location, description, and attendees
- Update existing events (time, location, attendees, description)
- Delete events
- Invite or remove attendees by providing an updated attendee list

### Contacts
- List and search your Outlook address book
- Create new contacts with name, email, phone, company, and job title

### Tasks (Microsoft To Do)
- List all task lists
- List tasks (filtered by status)
- Create tasks with due dates, notes, and priority

### Mailbox Settings
- Read your current settings (timezone, working hours, out-of-office status)
- Enable or disable out-of-office replies with custom internal and external messages
- Schedule out-of-office with start and end times

### OneDrive
- List files and folders (by path)
- Generate shareable links (view or edit, anonymous or organisation-scoped)

---

## How it works

```
Claude Desktop
      │
      │  stdio (JSON-RPC)
      ▼
proxy.mjs  ← local Node.js bridge script
      │
      │  HTTPS POST
      ▼
Cloudflare Worker  (outlook-mcp.bashar-basheer.workers.dev)
      │                    │
      │                    ├── Cloudflare KV  ← stores OAuth tokens
      │
      │  HTTPS (Bearer token)
      ▼
Microsoft Graph API  (graph.microsoft.com)
      │
      ▼
Your Outlook account
```

### Step by step

1. Claude Desktop spawns `proxy.mjs` as a local child process when it starts
2. Claude sends a JSON-RPC message over stdio to the proxy
3. The proxy forwards it as an HTTPS POST to the Cloudflare Worker's `/mcp` endpoint
4. The Worker looks up your OAuth access token from Cloudflare KV
5. If the token is within 5 minutes of expiring, it automatically refreshes it using the stored refresh token
6. The Worker calls the Microsoft Graph API on your behalf using the access token
7. The Graph API response is sanitised (HTML stripped, fields flattened) and returned to Claude
8. Claude reads the clean response and replies to you

---

## Where everything lives

### Cloudflare (cloud)
| Resource | Details |
|---|---|
| Worker | `outlook-mcp` — handles all requests |
| KV Namespace | `OAUTH_KV` (ID: `4994370056d545869e91491b115753fc`) — stores tokens |
| Worker URL | `https://outlook-mcp.bashar-basheer.workers.dev` |
| Dashboard | https://dash.cloudflare.com → Workers & Pages → outlook-mcp |

### Microsoft Azure (cloud)
| Resource | Details |
|---|---|
| App name | Outlook MCP Server |
| Client ID | `d36110b9-d68d-4e3c-b611-e31aaae38f94` |
| Tenant ID | `5fa1667c-fd27-4b4d-8a8a-bf5e9e38c626` |
| Redirect URI | `https://outlook-mcp.bashar-basheer.workers.dev/oauth/callback` |
| Portal | https://portal.azure.com → Microsoft Entra ID → App registrations |

> ⚠️ The Client Secret value is stored as an encrypted Wrangler secret in Cloudflare. It is not stored in any file. If it expires (set for 24 months), generate a new one in Azure and re-run `wrangler secret put MICROSOFT_CLIENT_SECRET` from the project folder.

> 🔒 The `/mcp` endpoint is protected by a shared secret (`MCP_SECRET`). Unauthenticated requests receive a 401. The secret is stored in Cloudflare encrypted secrets and passed to the local proxy via the `env` block in `claude_desktop_config.json`. To rotate it: generate a new value with `node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"`, run `wrangler secret put MCP_SECRET`, and update `claude_desktop_config.json`.

### Local machine
| File | Purpose |
|---|---|
| `outlook-mcp/src/index.ts` | Worker entry point — URL routing |
| `outlook-mcp/src/auth.ts` | OAuth token storage and refresh logic |
| `outlook-mcp/src/graph.ts` | Microsoft Graph API fetch wrapper |
| `outlook-mcp/src/tools.ts` | All 19 tool schemas and handler functions |
| `outlook-mcp/src/mcp.ts` | MCP JSON-RPC protocol dispatcher |
| `outlook-mcp/src/sanitize.ts` | Strips HTML and cleans Graph API responses |
| `outlook-mcp/src/types.ts` | TypeScript interfaces (Env, TokenData) |
| `outlook-mcp/wrangler.jsonc` | Worker config — KV binding, env vars |
| `outlook-mcp/proxy.mjs` | Local stdio ↔ HTTP bridge for Claude Desktop |
| `claude_desktop_config.json` | Claude Desktop MCP server registration |

---

## Permissions granted

The following Microsoft Graph API delegated permissions are active:

| Permission | What it allows |
|---|---|
| `Mail.ReadWrite` | Read and write all mail in all folders |
| `Mail.Send` | Send email |
| `Calendars.ReadWrite` | Full calendar access |
| `Files.ReadWrite.All` | Full OneDrive access |
| `Contacts.ReadWrite` | Read and write contacts |
| `MailboxSettings.ReadWrite` | Read and write mailbox settings including OOO |
| `Tasks.ReadWrite` | Read and write Microsoft To Do tasks |
| `Sites.Read.All` | Read SharePoint sites and OneDrive for Business |
| `offline_access` | Allows refresh tokens (stay authenticated) |
| `User.Read` | Read basic profile info |

---

## Maintenance

### Re-authenticate (if tokens expire or break)
Visit this URL in your browser and log in with your Microsoft account:
```
https://outlook-mcp.bashar-basheer.workers.dev/oauth/start
```

Check authentication status at any time:
```
https://outlook-mcp.bashar-basheer.workers.dev/oauth/status
```

### Deploy changes to the Worker
```bash
cd "/Users/basheerco/Documents/AI Agents & Workflows/Outlook-MCP/outlook-mcp"
wrangler deploy
```

### Check Worker logs (live)
```bash
cd "/Users/basheerco/Documents/AI Agents & Workflows/Outlook-MCP/outlook-mcp"
wrangler tail --format=pretty
```

### Rotate the Client Secret
1. Go to Azure portal → Microsoft Entra ID → App registrations → Outlook MCP Server → Certificates & secrets
2. Delete the old secret, create a new one, copy the value immediately
3. Run:
```bash
cd "/Users/basheerco/Documents/AI Agents & Workflows/Outlook-MCP/outlook-mcp"
wrangler secret put MICROSOFT_CLIENT_SECRET
```
4. Paste the new value when prompted — no redeploy needed

---

## How to add a new tool

All tools live in `src/tools.ts`. Adding one takes three steps:

**1. Add the schema** — in the `TOOLS` array, add a new object:
```typescript
{
  name: 'your_tool_name',
  description: 'What this tool does — Claude reads this to decide when to use it.',
  inputSchema: {
    type: 'object',
    properties: {
      param_name: { type: 'string', description: 'What this parameter is for' },
    },
    required: ['param_name'],
  },
},
```

**2. Add the handler** — add a `case` to the `callTool` switch and write the function:
```typescript
case 'your_tool_name': return yourToolHandler(token, args);
```

```typescript
async function yourToolHandler(token: string, args: ToolArgs): Promise<string> {
  const data = await graphRequest(token, '/me/some/graph/endpoint');
  return JSON.stringify(data);
}
```

**3. Deploy:**
```bash
wrangler deploy
```

The tool appears in Claude Desktop immediately after deploy — no restart required.

### Finding Graph API endpoints
Every tool calls the Microsoft Graph API. The full API reference is at:
https://learn.microsoft.com/en-us/graph/api/overview

Use **Graph Explorer** to test queries interactively before writing code:
https://developer.microsoft.com/en-us/graph/graph-explorer

---

## Key decisions and why

| Decision | Reason |
|---|---|
| Cloudflare Workers | Serverless, global, generous free tier, deploys in seconds |
| Cloudflare KV | Stateless Workers need somewhere to persist tokens between requests |
| OAuth 2.0 authorisation code flow | Required for delegated (act-on-behalf-of-user) Graph API access |
| Token auto-refresh | Access tokens expire after ~1 hour — refresh happens silently so tools never fail mid-use |
| Response sanitisation | Graph API returns verbose HTML-heavy responses — stripping them keeps Claude's context lean and responses fast |
| Local proxy (`proxy.mjs`) | Claude Desktop only supports stdio-based MCP servers — the proxy bridges stdio to the remote Worker over HTTPS |
| TypeScript | Type safety catches mistakes at build time, not at runtime when a tool call fails |

---

## Tags
`#tools` `#mcp` `#outlook` `#email` `#calendar` `#cloudflare` `#microsoft-graph` `#typescript`
