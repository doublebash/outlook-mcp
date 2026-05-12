import type { Env, TokenData } from './types';

// ── Single-flight refresh mutex contract ──────────────────────────────────────
// Microsoft Graph rolls refresh tokens: when /oauth2/v2.0/token returns a new
// refresh_token, the previously issued one is invalidated. If two callers race
// on an expired access token and both POST a refresh, the slower caller ends
// up holding a dead refresh_token and the user is forced through /oauth/start.
//
// We serialise refreshes via a KV mutex on `<TOKENS_KEY>:refresh_lock`. The
// lock value is a per-caller UUID so only the holder can release it. Cloudflare
// KV is last-write-wins — there is no native CAS — so after `put` we re-`get`
// and verify our UUID won. Losers wait briefly and re-read tokens; if the
// winner already refreshed, they reuse the fresh access token instead of
// triggering a second refresh. The lock auto-expires after 60s in case the
// holder dies mid-refresh. (60s is Cloudflare KV's minimum expirationTtl; a
// refresh normally completes in <5s, so the only effect of the floor is how
// long a dead holder blocks new refresh attempts.)

const TOKENS_KEY = 'oauth_tokens';
const REFRESH_LOCK_KEY = `${TOKENS_KEY}:refresh_lock`;
const LOCK_TTL_SECONDS = 60;
const WAIT_INTERVAL_MS = 200;
const MAX_WAIT_ATTEMPTS = 25; // 25 × 200ms ≈ 5s before giving up
const REFRESH_SKEW_MS = 5 * 60 * 1000; // refresh proactively within 5min of expiry

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

function isExpiringSoon(tokens: TokenData): boolean {
  return Date.now() >= tokens.expires_at - REFRESH_SKEW_MS;
}

// Returns a valid access token, refreshing automatically if needed
export async function getValidAccessToken(env: Env): Promise<string> {
  const tokens = await getTokens(env);
  if (!tokens) {
    throw new Error('Not authenticated. Visit /oauth/start in your browser to log in.');
  }

  if (!isExpiringSoon(tokens)) {
    return tokens.access_token;
  }

  return refreshWithLock(env);
}

// ── Refresh mutex ──────────────────────────────────────────────────────────────

function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Try to claim the refresh lock. Returns our lock token on success, null if
// another caller already owns it or beat us in the last-write-wins race.
async function acquireRefreshLock(env: Env): Promise<string | null> {
  const existing = await env.OAUTH_KV.get(REFRESH_LOCK_KEY);
  if (existing) return null;

  const myToken = crypto.randomUUID();
  await env.OAUTH_KV.put(REFRESH_LOCK_KEY, myToken, { expirationTtl: LOCK_TTL_SECONDS });

  // KV has no CAS — re-read and confirm we still own the lock.
  const winner = await env.OAUTH_KV.get(REFRESH_LOCK_KEY);
  return winner === myToken ? myToken : null;
}

// Release the lock only if we still hold it (TTL may have expired and another
// caller may have already taken over).
async function releaseRefreshLock(env: Env, lockToken: string): Promise<void> {
  const current = await env.OAUTH_KV.get(REFRESH_LOCK_KEY);
  if (current === lockToken) {
    await env.OAUTH_KV.delete(REFRESH_LOCK_KEY).catch(() => {});
  }
}

async function refreshWithLock(env: Env): Promise<string> {
  for (let attempt = 0; attempt < MAX_WAIT_ATTEMPTS; attempt++) {
    const lockToken = await acquireRefreshLock(env);

    if (lockToken) {
      try {
        // Re-read inside the lock — another caller may have refreshed between
        // our initial getTokens and lock acquisition. Always use the freshest
        // refresh_token from KV; using a stale one would invalidate the live one.
        const fresh = await getTokens(env);
        if (!fresh) {
          throw new Error('Not authenticated. Visit /oauth/start in your browser to log in.');
        }
        if (!isExpiringSoon(fresh)) {
          return fresh.access_token;
        }
        return await refreshAccessToken(env, fresh.refresh_token);
      } finally {
        await releaseRefreshLock(env, lockToken);
      }
    }

    // Lost the race / lock held by someone else. Wait, then check whether the
    // winner has already refreshed for us.
    await sleep(WAIT_INTERVAL_MS);
    const fresh = await getTokens(env);
    if (fresh && !isExpiringSoon(fresh)) {
      return fresh.access_token;
    }
  }

  // Don't fall through and refresh ourselves — that's exactly the race the
  // mutex exists to prevent. The lock auto-expires after 30s, so a retry
  // shortly will succeed.
  throw new Error(
    'Token refresh timed out waiting for another in-flight refresh. ' +
    'Retry shortly; if this persists, visit /oauth/start to re-authenticate.'
  );
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
