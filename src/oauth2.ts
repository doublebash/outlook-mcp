import type { Env } from './types';

// ── Types ──────────────────────────────────────────────────────────────────────

interface AuthCodeData {
  code_challenge: string;
  redirect_uri: string;
  client_id: string;
  state: string;
}

// ── Helpers ────────────────────────────────────────────────────────────────────

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

async function verifyPKCE(codeVerifier: string, codeChallenge: string): Promise<boolean> {
  try {
    const data = new TextEncoder().encode(codeVerifier);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = new Uint8Array(hashBuffer);
    let binary = '';
    for (let i = 0; i < hashArray.length; i++) {
      binary += String.fromCharCode(hashArray[i]);
    }
    const base64url = btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
    return base64url === codeChallenge;
  } catch {
    return false;
  }
}

// ── GET /.well-known/oauth-protected-resource (RFC 9396) ─────────────────────

export function oauthProtectedResource(env: Env): Response {
  return Response.json({
    resource: `${env.WORKER_URL}/mcp`,
    authorization_servers: [env.WORKER_URL],
  });
}

// ── GET /.well-known/oauth-authorization-server (RFC 8414) ───────────────────

export function oauthAuthorizationServer(env: Env): Response {
  return Response.json({
    issuer: env.WORKER_URL,
    authorization_endpoint: `${env.WORKER_URL}/authorize`,
    token_endpoint: `${env.WORKER_URL}/token`,
    registration_endpoint: `${env.WORKER_URL}/register`,
    response_types_supported: ['code'],
    code_challenge_methods_supported: ['S256'],
    grant_types_supported: ['authorization_code'],
    token_endpoint_auth_methods_supported: ['none'],
  });
}

// ── POST /register — dynamic client registration (RFC 7591) ──────────────────
// Claude Desktop requires this before starting the OAuth flow.
// We are a single-user server so we accept any registration and issue a client_id,
// but actual auth is always validated via MCP_SECRET.

export async function handleRegister(request: Request): Promise<Response> {
  let body: Record<string, unknown> = {};
  try {
    body = await request.json() as Record<string, unknown>;
  } catch {
    // proceed with empty body — registration metadata is optional
  }

  // Validate redirect_uris: must be strings and not plain HTTP
  // (HTTPS and custom URI schemes such as claude:// are allowed)
  const rawRedirectUris = body['redirect_uris'];
  if (rawRedirectUris !== undefined) {
    if (!Array.isArray(rawRedirectUris)) {
      return Response.json(
        { error: 'invalid_client_metadata', error_description: 'redirect_uris must be an array' },
        { status: 400 }
      );
    }
    for (const uri of rawRedirectUris) {
      if (!isAllowedRedirectUri(uri)) {
        return Response.json(
          { error: 'invalid_redirect_uri', error_description: 'redirect_uris must not use plain http://' },
          { status: 400 }
        );
      }
    }
  }

  // Only echo back fields we explicitly allow — don't reflect arbitrary input
  const redirectUris = Array.isArray(rawRedirectUris) ? rawRedirectUris as string[] : undefined;
  const clientName = typeof body['client_name'] === 'string' ? body['client_name'] : undefined;

  return Response.json(
    {
      client_id: crypto.randomUUID(),
      client_id_issued_at: Math.floor(Date.now() / 1000),
      ...(redirectUris ? { redirect_uris: redirectUris } : {}),
      ...(clientName ? { client_name: clientName } : {}),
      token_endpoint_auth_method: 'none',
      grant_types: ['authorization_code'],
      response_types: ['code'],
      code_challenge_methods: ['S256'],
    },
    { status: 201 }
  );
}

function isAllowedRedirectUri(uri: unknown): boolean {
  if (typeof uri !== 'string') return false;
  try {
    const parsed = new URL(uri);
    return parsed.protocol !== 'http:'; // Block plain HTTP; allow HTTPS and custom schemes
  } catch {
    return false; // Unparseable URI rejected
  }
}

// ── GET /authorize — approval page ────────────────────────────────────────────

export function handleAuthorize(request: Request): Response {
  const url = new URL(request.url);
  const p = url.searchParams;

  const clientId            = p.get('client_id') ?? '';
  const redirectUri         = p.get('redirect_uri') ?? '';
  const state               = p.get('state') ?? '';
  const codeChallenge       = p.get('code_challenge') ?? '';
  const codeChallengeMethod = p.get('code_challenge_method') ?? 'S256';
  const showError           = p.get('error') === 'invalid_code';

  const errorBanner = showError
    ? `<div class="error">Incorrect access code — please try again.</div>`
    : '';

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Authorise — Outlook MCP</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: #f0f2f5;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 1rem;
    }
    .card {
      background: #fff;
      border-radius: 14px;
      padding: 2.5rem 2rem;
      max-width: 400px;
      width: 100%;
      box-shadow: 0 4px 32px rgba(0,0,0,0.09);
    }
    .logo { font-size: 2.25rem; margin-bottom: 1.25rem; }
    h1 { font-size: 1.3rem; font-weight: 650; color: #111; margin-bottom: 0.5rem; }
    .subtitle { font-size: 0.9rem; color: #555; line-height: 1.55; margin-bottom: 1.75rem; }
    .error {
      font-size: 0.875rem;
      color: #b91c1c;
      background: #fef2f2;
      border: 1px solid #fecaca;
      border-radius: 8px;
      padding: 0.625rem 0.875rem;
      margin-bottom: 1.25rem;
    }
    label { display: block; font-size: 0.8125rem; font-weight: 600; color: #333; margin-bottom: 0.375rem; }
    input[type="password"] {
      width: 100%;
      padding: 0.625rem 0.875rem;
      border: 1.5px solid #d1d5db;
      border-radius: 8px;
      font-size: 0.9375rem;
      outline: none;
      transition: border-color 0.15s;
    }
    input[type="password"]:focus { border-color: #0078d4; box-shadow: 0 0 0 3px rgba(0,120,212,0.12); }
    button {
      width: 100%;
      margin-top: 1.125rem;
      padding: 0.75rem;
      background: #0078d4;
      color: #fff;
      border: none;
      border-radius: 8px;
      font-size: 0.9375rem;
      font-weight: 600;
      cursor: pointer;
      transition: background 0.15s;
    }
    button:hover { background: #106ebe; }
    .hint { font-size: 0.8rem; color: #888; margin-top: 1.25rem; line-height: 1.5; }
    code { font-size: 0.78rem; background: #f3f4f6; border-radius: 4px; padding: 0.1em 0.35em; }
  </style>
</head>
<body>
  <div class="card">
    <div class="logo">📬</div>
    <h1>Authorise Outlook MCP</h1>
    <p class="subtitle">
      Claude.ai is requesting access to your Outlook email, calendar, contacts,
      tasks, and OneDrive through your personal MCP server.
    </p>
    ${errorBanner}
    <form method="POST" action="/approve">
      <input type="hidden" name="client_id"             value="${escapeHtml(clientId)}">
      <input type="hidden" name="redirect_uri"          value="${escapeHtml(redirectUri)}">
      <input type="hidden" name="state"                 value="${escapeHtml(state)}">
      <input type="hidden" name="code_challenge"        value="${escapeHtml(codeChallenge)}">
      <input type="hidden" name="code_challenge_method" value="${escapeHtml(codeChallengeMethod)}">
      <label for="access_code">Access code</label>
      <input
        type="password"
        id="access_code"
        name="access_code"
        placeholder="Paste your MCP_SECRET"
        autocomplete="off"
        required
      >
      <button type="submit">Authorise access</button>
    </form>
    <p class="hint">
      Find your access code in
      <code>~/Library/Application Support/Claude/claude_desktop_config.json</code>
      under the <code>outlook</code> entry's <code>MCP_SECRET</code> env value.
    </p>
  </div>
</body>
</html>`;

  return new Response(html, {
    headers: { 'Content-Type': 'text/html; charset=utf-8' },
  });
}

// ── POST /approve — validate code, issue auth code, redirect ─────────────────

export async function handleApprove(request: Request, env: Env): Promise<Response> {
  let form: FormData;
  try {
    form = await request.formData();
  } catch {
    return new Response('Bad request', { status: 400 });
  }

  const accessCode      = (form.get('access_code') as string | null) ?? '';
  const codeChallenge   = (form.get('code_challenge') as string | null) ?? '';
  const redirectUri     = (form.get('redirect_uri') as string | null) ?? '';
  const state           = (form.get('state') as string | null) ?? '';
  const clientId        = (form.get('client_id') as string | null) ?? '';
  const codeChallengeMethod = (form.get('code_challenge_method') as string | null) ?? 'S256';

  // Validate submitted access code against MCP_SECRET
  if (!accessCode || accessCode !== env.MCP_SECRET) {
    const params = new URLSearchParams({
      client_id: clientId,
      redirect_uri: redirectUri,
      state,
      code_challenge: codeChallenge,
      code_challenge_method: codeChallengeMethod,
      error: 'invalid_code',
    });
    return Response.redirect(`${env.WORKER_URL}/authorize?${params}`, 302);
  }

  // Generate a one-time auth code and store it in KV (TTL: 5 minutes)
  const authCode = crypto.randomUUID();
  const codeData: AuthCodeData = { code_challenge: codeChallenge, redirect_uri: redirectUri, client_id: clientId, state };

  try {
    await env.OUTLOOK_OAUTH_KV.put(authCode, JSON.stringify(codeData), { expirationTtl: 300 });
  } catch {
    return new Response('Internal server error', { status: 500 });
  }

  // Redirect back to the client with the auth code
  const callbackUrl = new URL(redirectUri);
  callbackUrl.searchParams.set('code', authCode);
  if (state) callbackUrl.searchParams.set('state', state);

  return Response.redirect(callbackUrl.toString(), 302);
}

// ── POST /token — exchange code for access token ──────────────────────────────

export async function handleToken(request: Request, env: Env): Promise<Response> {
  let body: URLSearchParams;
  try {
    const text = await request.text();
    body = new URLSearchParams(text);
  } catch {
    return Response.json({ error: 'invalid_request' }, { status: 400 });
  }

  const grantType = body.get('grant_type');

  // ── authorization_code ────────────────────────────────────────────────────
  if (grantType === 'authorization_code') {
    const code         = body.get('code');
    const codeVerifier = body.get('code_verifier');
    const redirectUri  = body.get('redirect_uri');

    if (!code || !codeVerifier || !redirectUri) {
      return Response.json({ error: 'invalid_request' }, { status: 400 });
    }

    // Look up the auth code
    let codeData: AuthCodeData | null = null;
    try {
      const raw = await env.OUTLOOK_OAUTH_KV.get(code);
      if (raw) codeData = JSON.parse(raw) as AuthCodeData;
    } catch {
      return Response.json({ error: 'server_error' }, { status: 500 });
    }

    if (!codeData) {
      return Response.json({ error: 'invalid_grant' }, { status: 400 });
    }

    // Verify redirect_uri matches exactly
    if (codeData.redirect_uri !== redirectUri) {
      return Response.json({ error: 'invalid_grant' }, { status: 400 });
    }

    // Verify PKCE: SHA-256(code_verifier) base64url must equal stored code_challenge
    const pkceValid = await verifyPKCE(codeVerifier, codeData.code_challenge);
    if (!pkceValid) {
      return Response.json({ error: 'invalid_grant' }, { status: 400 });
    }

    // Consume the code — one-time use
    try {
      await env.OUTLOOK_OAUTH_KV.delete(code);
    } catch {
      // Non-fatal: code expires in 5 minutes anyway
    }

    return Response.json({
      access_token: env.MCP_SECRET,
      token_type: 'Bearer',
      expires_in: 7776000,
    });
  }

  return Response.json({ error: 'unsupported_grant_type' }, { status: 400 });
}
