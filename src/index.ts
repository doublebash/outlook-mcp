import { SCOPES, getTokens, saveTokens } from './auth';
import { handleMCP } from './mcp';
import {
  oauthProtectedResource,
  oauthAuthorizationServer,
  handleAuthorize,
  handleApprove,
  handleToken,
  handleRegister,
} from './oauth2';
import type { Env, TokenData } from './types';

// ── CORS ───────────────────────────────────────────────────────────────────────

const CORS_HEADERS: Record<string, string> = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept, Mcp-Session-Id',
  'Access-Control-Expose-Headers': 'WWW-Authenticate',
};

function withCors(response: Response): Response {
  const newResponse = new Response(response.body, response);
  for (const [key, value] of Object.entries(CORS_HEADERS)) {
    newResponse.headers.set(key, value);
  }
  return newResponse;
}

// ── HTML escaping ──────────────────────────────────────────────────────────────

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// ── Router ─────────────────────────────────────────────────────────────────────

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);
    const { pathname } = url;
    const method = request.method;

    // CORS preflight
    if (method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    // ── OAuth 2.0 public endpoints (no auth required) ──
    if (method === 'GET'  && pathname === '/.well-known/oauth-protected-resource')  return withCors(oauthProtectedResource(env));
    if (method === 'GET'  && pathname === '/.well-known/oauth-authorization-server') return withCors(oauthAuthorizationServer(env));
    if (method === 'GET'  && pathname === '/authorize')                              return withCors(handleAuthorize(request));
    if (method === 'POST' && pathname === '/approve')                                return withCors(await handleApprove(request, env));
    if (method === 'POST' && pathname === '/token')                                  return withCors(await handleToken(request, env));
    if (method === 'POST' && pathname === '/register')                               return withCors(await handleRegister(request));

    // ── Microsoft OAuth (one-time setup — visit /oauth/start in browser) ──
    if (method === 'GET'  && pathname === '/oauth/start')    return withCors(oauthStartForm());
    if (method === 'POST' && pathname === '/oauth/start')    return withCors(await oauthStart(request, env));
    if (method === 'GET'  && pathname === '/oauth/callback') return withCors(await oauthCallback(request, env));
    if (method === 'GET'  && pathname === '/oauth/status')   return withCors(await oauthStatus(request, env));

    // ── MCP endpoint ──
    if (method === 'POST' && pathname === '/mcp') return withCors(await handleMCP(request, env));

    return withCors(new Response('Not found', { status: 404 }));
  },
} satisfies ExportedHandler<Env>;

// ── Microsoft OAuth: Step 1a — show setup form (GET) ─────────────────────────
// The secret is submitted via POST body so it never appears in the URL,
// browser history, or server logs.

function oauthStartForm(error = false): Response {
  const errorBanner = error
    ? `<div class="error">Incorrect access code — please try again.</div>`
    : '';

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Connect Outlook</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: #f0f2f5; min-height: 100vh;
      display: flex; align-items: center; justify-content: center; padding: 1rem;
    }
    .card {
      background: #fff; border-radius: 14px; padding: 2.5rem 2rem;
      max-width: 400px; width: 100%; box-shadow: 0 4px 32px rgba(0,0,0,0.09);
    }
    .logo { font-size: 2.25rem; margin-bottom: 1.25rem; }
    h1 { font-size: 1.3rem; font-weight: 650; color: #111; margin-bottom: 0.5rem; }
    p { font-size: 0.9rem; color: #555; line-height: 1.55; margin-bottom: 1.75rem; }
    .error {
      font-size: 0.875rem; color: #b91c1c; background: #fef2f2;
      border: 1px solid #fecaca; border-radius: 8px;
      padding: 0.625rem 0.875rem; margin-bottom: 1.25rem;
    }
    label { display: block; font-size: 0.8125rem; font-weight: 600; color: #333; margin-bottom: 0.375rem; }
    input[type="password"] {
      width: 100%; padding: 0.625rem 0.875rem;
      border: 1.5px solid #d1d5db; border-radius: 8px;
      font-size: 0.9375rem; outline: none; transition: border-color 0.15s;
    }
    input[type="password"]:focus { border-color: #0078d4; box-shadow: 0 0 0 3px rgba(0,120,212,0.12); }
    button {
      width: 100%; margin-top: 1.125rem; padding: 0.75rem;
      background: #0078d4; color: #fff; border: none; border-radius: 8px;
      font-size: 0.9375rem; font-weight: 600; cursor: pointer; transition: background 0.15s;
    }
    button:hover { background: #106ebe; }
  </style>
</head>
<body>
  <div class="card">
    <div class="logo">📬</div>
    <h1>Connect Outlook</h1>
    <p>Enter your MCP_SECRET to start the Microsoft authorisation flow.</p>
    ${errorBanner}
    <form method="POST" action="/oauth/start">
      <label for="secret">Access code</label>
      <input type="password" id="secret" name="secret"
        placeholder="Paste your MCP_SECRET" autocomplete="off" required>
      <button type="submit">Connect to Microsoft →</button>
    </form>
  </div>
</body>
</html>`;

  return new Response(html, { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
}

// ── Microsoft OAuth: Step 1b — validate secret and redirect to Microsoft (POST) ──

async function oauthStart(request: Request, env: Env): Promise<Response> {
  let form: FormData;
  try {
    form = await request.formData();
  } catch {
    return new Response('Bad request', { status: 400 });
  }

  const secret = (form.get('secret') as string | null) ?? '';
  if (secret !== env.MCP_SECRET) {
    return oauthStartForm(true); // Show form again with error banner
  }

  const params = new URLSearchParams({
    client_id: env.MICROSOFT_CLIENT_ID,
    response_type: 'code',
    redirect_uri: `${env.WORKER_URL}/oauth/callback`,
    scope: SCOPES,
    response_mode: 'query',
  });

  const authUrl =
    `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/authorize?${params}`;

  return Response.redirect(authUrl, 302);
}

// ── Microsoft OAuth: Step 2 — handle callback ─────────────────────────────────

async function oauthCallback(request: Request, env: Env): Promise<Response> {
  const url = new URL(request.url);
  const code = url.searchParams.get('code');
  const errorParam = url.searchParams.get('error');

  if (errorParam) {
    // Don't reflect error_description — it comes from Microsoft and may contain details.
    return new Response(`OAuth error: ${escapeHtml(errorParam)}`, { status: 400 });
  }

  if (!code) {
    return new Response('Missing authorization code', { status: 400 });
  }

  const response = await fetch(
    `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: env.MICROSOFT_CLIENT_ID,
        client_secret: env.MICROSOFT_CLIENT_SECRET,
        grant_type: 'authorization_code',
        code,
        redirect_uri: `${env.WORKER_URL}/oauth/callback`,
        scope: SCOPES,
      }),
    }
  );

  if (!response.ok) {
    return new Response('Token exchange failed. Please try again.', { status: 500 });
  }

  const data = await response.json() as {
    access_token: string;
    refresh_token: string;
    expires_in: number;
  };

  const tokens: TokenData = {
    access_token: data.access_token,
    refresh_token: data.refresh_token,
    expires_at: Date.now() + data.expires_in * 1000,
  };

  await saveTokens(env, tokens);

  return new Response(
    `<!DOCTYPE html>
    <html>
      <head><title>Authenticated</title></head>
      <body style="font-family:sans-serif;padding:2rem;max-width:500px;margin:auto">
        <h2>Authentication successful</h2>
        <p>Your Outlook MCP server is now connected to Microsoft. You can close this window.</p>
      </body>
    </html>`,
    { headers: { 'Content-Type': 'text/html' } }
  );
}

// ── Microsoft OAuth: Status check ─────────────────────────────────────────────
// Requires Bearer token to prevent unauthenticated enumeration.

async function oauthStatus(request: Request, env: Env): Promise<Response> {
  const authHeader = request.headers.get('Authorization');
  const bearerToken = authHeader?.startsWith('Bearer ') ? authHeader.slice(7) : null;

  const authorised = bearerToken !== null && bearerToken === env.MCP_SECRET;

  if (!authorised) {
    return new Response('Unauthorised', { status: 401 });
  }

  const tokens = await getTokens(env);

  if (!tokens) {
    return Response.json({
      authenticated: false,
      message: 'No tokens found. Visit /oauth/start to authenticate.',
    });
  }

  const now = Date.now();
  const expired = now >= tokens.expires_at;
  const minutesRemaining = Math.round((tokens.expires_at - now) / 60000);

  return Response.json({
    authenticated: true,
    expired,
    expires_at: new Date(tokens.expires_at).toISOString(),
    minutes_remaining: expired ? 0 : minutesRemaining,
  });
}
