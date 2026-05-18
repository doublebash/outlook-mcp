import { Hono } from "hono";
import {
	createBearerMiddleware,
	createCors,
	createMcpRouter,
	createOAuthServer,
	createRateLimit,
	DEFAULT_CLAUDE_ORIGINS,
	OAUTH_PUBLIC_PATHS,
	escapeHtml,
	timingSafeEqual,
} from "@bashco/mcp-toolkit";
import { SCOPES, clearTokens, getTokens, saveTokens } from "./auth.js";
import { dispatchToolCall, toolDefinitions } from "./tools/index.js";
import type { Env, OutlookTokenData } from "./types.js";

const SERVER_NAME = "outlook-mcp";
const SERVER_VERSION = "2.0.0";
const SUPPORTED_PROTOCOL_VERSIONS = ["2024-11-05", "2025-03-26"] as const;
const DEFAULT_PROTOCOL_VERSION: (typeof SUPPORTED_PROTOCOL_VERSIONS)[number] =
	"2024-11-05";

// Claude.ai's web/desktop OAuth client registers redirect URIs under these hosts.
const ALLOWED_REDIRECT_HOSTS = new Set([
	"claude.ai",
	"api.claude.ai",
	"claude.com",
	"api.claude.com",
]);

// 15 MB MCP body cap — generous enough to accept tool calls with 10 MB of
// base64-encoded attachments (which inflate ~33% on the wire).
const MCP_BODY_CAP_BYTES = 15 * 1024 * 1024;

const PUBLIC_PATHS = new Set<string>([
	...OAUTH_PUBLIC_PATHS,
	"/oauth/start",
	"/oauth/callback",
]);

const app = new Hono<{ Bindings: Env }>();

app.use("*", createCors({ allowedOrigins: DEFAULT_CLAUDE_ORIGINS }));
app.use(
	"*",
	createBearerMiddleware({
		kv: (env) => (env as Env).OAUTH_KV,
		publicPaths: PUBLIC_PATHS,
		realm: SERVER_NAME,
	}),
);

const oauth = createOAuthServer<Env>({
	serverName: SERVER_NAME,
	serverDescription:
		"Claude is requesting access to your Outlook account (email, calendar, contacts, tasks, OneDrive, mailbox settings). Enter the operator-only approval code to continue.",
	logo: "📬",
	approvalCodeName: "MCP_APPROVAL_CODE",
	kv: (env) => env.OAUTH_KV,
	approvalCodeSecret: (env) => env.MCP_APPROVAL_CODE,
	allowedRedirectHosts: ALLOWED_REDIRECT_HOSTS,
	rateLimiters: {
		approve: createRateLimit<Env>({
			binding: (env) => env.RATE_LIMIT_APPROVE,
			bucketName: "approve",
		}),
		token: createRateLimit<Env>({
			binding: (env) => env.RATE_LIMIT_TOKEN,
			bucketName: "token",
		}),
		register: createRateLimit<Env>({
			binding: (env) => env.RATE_LIMIT_REGISTER,
			bucketName: "register",
		}),
	},
});

const mcp = createMcpRouter<Env>({
	serverName: SERVER_NAME,
	serverVersion: SERVER_VERSION,
	protocolVersions: SUPPORTED_PROTOCOL_VERSIONS,
	defaultProtocolVersion: DEFAULT_PROTOCOL_VERSION,
	toolDefinitions,
	dispatch: async (env, _ctx, name, args) => dispatchToolCall(env, name, args),
	rateLimiter: createRateLimit<Env>({
		binding: (env) => env.RATE_LIMIT_MCP,
		bucketName: "mcp",
	}),
	maxBodyBytes: MCP_BODY_CAP_BYTES,
});

// ── Microsoft upstream OAuth (operator-facing one-time setup) ────────────────

const oauthStartRateLimit = createRateLimit<Env>({
	binding: (env) => env.RATE_LIMIT_OAUTH_START,
	bucketName: "oauth_start",
});

app.get("/oauth/start", (c) => c.html(oauthStartForm()));

app.post("/oauth/start", oauthStartRateLimit, async (c) => {
	let form: FormData;
	try {
		form = await c.req.formData();
	} catch {
		return c.text("Bad request", 400);
	}
	const secret = (form.get("secret") as string | null) ?? "";
	if (!secret || !(await timingSafeEqual(secret, c.env.MCP_APPROVAL_CODE ?? ""))) {
		return c.html(oauthStartForm(true), 401);
	}

	// CSRF state cookie — generated, sent to Microsoft, verified on /oauth/callback.
	const state = crypto.randomUUID();
	const params = new URLSearchParams({
		client_id: c.env.MICROSOFT_CLIENT_ID,
		response_type: "code",
		redirect_uri: `${c.env.WORKER_URL}/oauth/callback`,
		scope: SCOPES,
		response_mode: "query",
		state,
	});

	const authUrl = `https://login.microsoftonline.com/${c.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/authorize?${params}`;
	return new Response(null, {
		status: 302,
		headers: {
			Location: authUrl,
			"Set-Cookie": `outlook_state=${state}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=300`,
		},
	});
});

app.get("/oauth/callback", async (c) => {
	const url = new URL(c.req.url);
	const code = url.searchParams.get("code");
	const returnedState = url.searchParams.get("state");
	const errorParam = url.searchParams.get("error");

	if (errorParam) {
		// Don't reflect error_description — it comes from Microsoft and may contain details.
		return c.text(`OAuth error: ${escapeHtml(errorParam)}`, 400);
	}
	if (!code) return c.text("Missing authorisation code", 400);

	const cookieHeader = c.req.header("Cookie") ?? "";
	const stateCookie = cookieHeader
		.split(";")
		.map((s) => s.trim())
		.find((s) => s.startsWith("outlook_state="))
		?.split("=")
		.slice(1)
		.join("=");
	if (!stateCookie || stateCookie !== returnedState) {
		return c.text("Invalid state parameter — possible CSRF attack. Please try again.", 400);
	}

	const response = await fetch(
		`https://login.microsoftonline.com/${c.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
		{
			method: "POST",
			headers: { "Content-Type": "application/x-www-form-urlencoded" },
			body: new URLSearchParams({
				client_id: c.env.MICROSOFT_CLIENT_ID,
				client_secret: c.env.MICROSOFT_CLIENT_SECRET,
				grant_type: "authorization_code",
				code,
				redirect_uri: `${c.env.WORKER_URL}/oauth/callback`,
				scope: SCOPES,
			}),
		},
	);

	if (!response.ok) {
		return c.text("Token exchange failed. Please try again.", 500);
	}

	const data = (await response.json()) as {
		access_token: string;
		refresh_token: string;
		expires_in: number;
	};

	const tokens: OutlookTokenData = {
		access_token: data.access_token,
		refresh_token: data.refresh_token,
		expires_at: Date.now() + data.expires_in * 1000,
	};
	await saveTokens(c.env, tokens);

	return new Response(
		`<!DOCTYPE html>
<html><head><title>Connected</title></head>
<body style="font-family:sans-serif;padding:2rem;max-width:500px;margin:auto">
  <h2>Outlook connected successfully</h2>
  <p>Your Outlook MCP server is now connected to Microsoft. You can close this window and start using it in Claude.</p>
</body></html>`,
		{
			headers: {
				"Content-Type": "text/html; charset=utf-8",
				"Set-Cookie": "outlook_state=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0",
			},
		},
	);
});

app.get("/oauth/status", async (c) => {
	// Bearer middleware has already validated the per-client token.
	const tokens = await getTokens(c.env);
	if (!tokens) {
		return c.json({
			authenticated: false,
			message:
				"Not connected. Visit /oauth/start to authorise Outlook (paste MCP_APPROVAL_CODE when asked).",
		});
	}
	const now = Date.now();
	const expired = now >= tokens.expires_at;
	const minutesRemaining = Math.round((tokens.expires_at - now) / 60000);
	return c.json({
		authenticated: true,
		expired,
		expires_at: new Date(tokens.expires_at).toISOString(),
		minutes_remaining: expired ? 0 : minutesRemaining,
	});
});

// Optional: revoke endpoint for explicit upstream-token clearing.
// Bearer-gated so anyone holding a valid Claude.ai connection can clear.
app.post("/oauth/disconnect", async (c) => {
	await clearTokens(c.env);
	return c.json({
		success: true,
		message: "Outlook tokens cleared. Visit /oauth/start to reconnect.",
	});
});

app.route("/", oauth.routes);
app.route("/", mcp);

export default app;

// ── /oauth/start form ─────────────────────────────────────────────────────────

function oauthStartForm(error = false): string {
	const errorBanner = error
		? `<div class="error">Incorrect approval code — please try again.</div>`
		: "";
	return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Connect Outlook</title>
  <style>
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#f0f2f5;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:1rem}
    .card{background:#fff;border-radius:14px;padding:2.5rem 2rem;max-width:400px;width:100%;box-shadow:0 4px 32px rgba(0,0,0,0.09)}
    .logo{font-size:2.25rem;margin-bottom:1.25rem}
    h1{font-size:1.3rem;font-weight:650;color:#111;margin-bottom:0.5rem}
    p{font-size:0.9rem;color:#555;line-height:1.55;margin-bottom:1.75rem}
    .error{font-size:0.875rem;color:#b91c1c;background:#fef2f2;border:1px solid #fecaca;border-radius:8px;padding:0.625rem 0.875rem;margin-bottom:1.25rem}
    label{display:block;font-size:0.8125rem;font-weight:600;color:#333;margin-bottom:0.375rem}
    input[type="password"]{width:100%;padding:0.625rem 0.875rem;border:1.5px solid #d1d5db;border-radius:8px;font-size:0.9375rem;outline:none;transition:border-color 0.15s}
    input[type="password"]:focus{border-color:#0078d4;box-shadow:0 0 0 3px rgba(0,120,212,0.12)}
    button{width:100%;margin-top:1.125rem;padding:0.75rem;background:#0078d4;color:#fff;border:none;border-radius:8px;font-size:0.9375rem;font-weight:600;cursor:pointer;transition:background 0.15s}
    button:hover{background:#106ebe}
    code{font-size:0.78rem;background:#eef;border-radius:4px;padding:0.1em 0.35em}
  </style>
</head>
<body>
  <div class="card">
    <div class="logo">📬</div>
    <h1>Connect Outlook</h1>
    <p>Enter your <code>MCP_APPROVAL_CODE</code> to start the Microsoft authorisation flow. This is the operator-only secret set via <code>wrangler secret put</code> — it is NOT used as a runtime bearer.</p>
    ${errorBanner}
    <form method="POST" action="/oauth/start" autocomplete="off">
      <label for="secret">Approval code</label>
      <input type="password" id="secret" name="secret" autocomplete="one-time-code" required autofocus>
      <button type="submit">Connect to Microsoft →</button>
    </form>
  </div>
</body>
</html>`;
}
