import type { Env, TokenData } from './types';

const TOKENS_KEY = 'oauth_tokens';

// ── Token encryption (AES-256-GCM) ────────────────────────────────────────────
// Tokens are encrypted at rest in KV so a KV namespace compromise does not
// expose live Microsoft credentials. The key is derived from MCP_SECRET via HKDF.
//
// ⚠️  If MCP_SECRET changes, stored tokens cannot be decrypted and the user
//     must re-authenticate via /oauth/start.

async function deriveKey(secret: string): Promise<CryptoKey> {
  const keyMaterial = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(secret),
    'HKDF',
    false,
    ['deriveKey']
  );
  return crypto.subtle.deriveKey(
    {
      name: 'HKDF',
      hash: 'SHA-256',
      salt: new TextEncoder().encode('outlook-mcp-v1'),
      info: new TextEncoder().encode('token-storage'),
    },
    keyMaterial,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
}

async function encryptData(plaintext: string, secret: string): Promise<string> {
  const key = await deriveKey(secret);
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ciphertext = await crypto.subtle.encrypt(
    { name: 'AES-GCM', iv },
    key,
    new TextEncoder().encode(plaintext)
  );
  const combined = new Uint8Array(iv.length + ciphertext.byteLength);
  combined.set(iv);
  combined.set(new Uint8Array(ciphertext), iv.length);
  let binary = '';
  for (let i = 0; i < combined.length; i++) binary += String.fromCharCode(combined[i]!);
  return btoa(binary);
}

async function decryptData(encrypted: string, secret: string): Promise<string> {
  const key = await deriveKey(secret);
  const combined = Uint8Array.from(atob(encrypted), c => c.charCodeAt(0));
  const iv = combined.slice(0, 12);
  const ciphertext = combined.slice(12);
  const decrypted = await crypto.subtle.decrypt(
    { name: 'AES-GCM', iv },
    key,
    ciphertext
  );
  return new TextDecoder().decode(decrypted);
}

// ── Scopes ─────────────────────────────────────────────────────────────────────

// Every permission scope we requested in Azure
export const SCOPES = [
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.ReadWrite',
  'Files.ReadWrite.All',
  'Contacts.ReadWrite',
  'MailboxSettings.ReadWrite',
  'Tasks.ReadWrite',
  'Sites.Read.All',
  'offline_access',
  'User.Read',
].join(' ');

// ── Token storage ──────────────────────────────────────────────────────────────

export async function getTokens(env: Env): Promise<TokenData | null> {
  const raw = await env.OAUTH_KV.get(TOKENS_KEY);
  if (!raw) return null;
  try {
    const decrypted = await decryptData(raw, env.MCP_SECRET);
    return JSON.parse(decrypted) as TokenData;
  } catch {
    // Decryption failed — likely plaintext tokens from before this fix,
    // or MCP_SECRET was rotated. Clear them so the user re-authenticates cleanly.
    await env.OAUTH_KV.delete(TOKENS_KEY).catch(() => {});
    return null;
  }
}

export async function saveTokens(env: Env, tokens: TokenData): Promise<void> {
  const encrypted = await encryptData(JSON.stringify(tokens), env.MCP_SECRET);
  await env.OAUTH_KV.put(TOKENS_KEY, encrypted);
}

// ── Access token helpers ───────────────────────────────────────────────────────

// Returns a valid access token, refreshing automatically if needed
export async function getValidAccessToken(env: Env): Promise<string> {
  const tokens = await getTokens(env);
  if (!tokens) {
    throw new Error('Not authenticated. Visit /oauth/start in your browser to log in.');
  }

  // Refresh proactively if within 5 minutes of expiry
  const fiveMinutes = 5 * 60 * 1000;
  if (Date.now() >= tokens.expires_at - fiveMinutes) {
    return refreshAccessToken(env, tokens.refresh_token);
  }

  return tokens.access_token;
}

// ── Token refresh ──────────────────────────────────────────────────────────────

async function refreshAccessToken(env: Env, refreshToken: string): Promise<string> {
  const response = await fetch(
    `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: env.MICROSOFT_CLIENT_ID,
        client_secret: env.MICROSOFT_CLIENT_SECRET,
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
        scope: SCOPES,
      }),
    }
  );

  if (!response.ok) {
    // Don't include raw Microsoft response body — it may contain diagnostic details.
    throw new Error(
      `Token refresh failed (HTTP ${response.status}). ` +
      `Visit /oauth/start to re-authenticate.`
    );
  }

  const data = await response.json() as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
  };

  const updated: TokenData = {
    access_token: data.access_token,
    // Microsoft doesn't always return a new refresh token — keep the old one if absent
    refresh_token: data.refresh_token ?? refreshToken,
    expires_at: Date.now() + data.expires_in * 1000,
  };

  await saveTokens(env, updated);
  return updated.access_token;
}
