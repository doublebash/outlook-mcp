import { createProxiedOAuthStore, type ProxiedOAuthStore } from "@bashco/mcp-toolkit";
import type { Env, OutlookTokenData } from "./types.js";

const TOKENS_KEY = "oauth_tokens";
const ENCRYPTION_NAMESPACE = "outlook-mcp-v2";

// Plaintext breadcrumb written next to the encrypted tokens. Holds no secret
// material — only "a connection was established, under this namespace, at this
// time".
//
// Why it exists: the toolkit's store treats an undecryptable blob as absent.
// `getTokens()` catches the decrypt error, DELETES the KV entry, and returns
// null, so `getValidAccessToken()` then fails with the same "not authenticated"
// as a mailbox that was never connected — and the ciphertext is gone before
// anyone can diagnose it. That is the silent-failure shape that hid the
// bashco-dashboard Xero break for weeks.
//
// The marker survives that deletion, so a failure can be attributed: marker
// present + tokens gone means the stored tokens stopped decrypting (rotated
// MCP_APPROVAL_CODE, or upstream crypto drift), not "never connected".
const CONNECTION_MARKER_KEY = "oauth_tokens:connected";

interface ConnectionMarker {
	connected_at: string;
	namespace: string;
}

// Every Microsoft Graph delegated permission we ask for.
// `offline_access` is required for refresh tokens.
// `OnlineMeetingRecording.Read.All` and `OnlineMeetingTranscript.Read.All`
// require admin consent in Azure AD even on single-tenant setups.
export const SCOPES = [
	"Mail.ReadWrite",
	"Mail.Send",
	"Calendars.ReadWrite",
	"Files.ReadWrite.All",
	"Contacts.ReadWrite",
	"MailboxSettings.ReadWrite",
	"Tasks.ReadWrite",
	"Sites.Read.All",
	"OnlineMeetings.Read",
	"OnlineMeetingRecording.Read.All",
	"OnlineMeetingTranscript.Read.All",
	"offline_access",
	"User.Read",
].join(" ");

function makeStore(env: Env): ProxiedOAuthStore {
	return createProxiedOAuthStore({
		kv: env.OAUTH_KV,
		tokensKey: TOKENS_KEY,
		// HKDF secret derives from the operator approval code — never from a
		// runtime bearer. Rotating MCP_APPROVAL_CODE invalidates stored tokens
		// and forces a clean re-auth via /oauth/start.
		encryptionSecret: env.MCP_APPROVAL_CODE,
		encryptionNamespace: ENCRYPTION_NAMESPACE,
		refreshEndpoint: `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
		refreshBody: (refreshToken) => ({
			client_id: env.MICROSOFT_CLIENT_ID,
			client_secret: env.MICROSOFT_CLIENT_SECRET,
			grant_type: "refresh_token",
			refresh_token: refreshToken,
			scope: SCOPES,
		}),
		// Microsoft does not expose a standalone refresh-token revocation endpoint
		// for personal accounts. The closest thing is account-level consent removal
		// via myapps.microsoft.com. Best-effort: hit the v2 logout endpoint which
		// invalidates the session for some flows. Full revocation still requires
		// the user revoke the app in their Microsoft account.
	});
}

export async function getTokens(env: Env): Promise<OutlookTokenData | null> {
	const raw = await makeStore(env).getTokens();
	if (!raw) return null;
	return raw as unknown as OutlookTokenData;
}

export async function saveTokens(env: Env, tokens: OutlookTokenData): Promise<void> {
	await makeStore(env).saveTokens(
		tokens as unknown as Parameters<ProxiedOAuthStore["saveTokens"]>[0],
	);
	const marker: ConnectionMarker = {
		connected_at: new Date().toISOString(),
		namespace: ENCRYPTION_NAMESPACE,
	};
	// Best-effort: the marker is diagnostic only, so a write failure must never
	// break an otherwise-successful connect.
	await env.OAUTH_KV.put(CONNECTION_MARKER_KEY, JSON.stringify(marker)).catch(() => {});
}

async function readConnectionMarker(env: Env): Promise<ConnectionMarker | null> {
	try {
		const raw = await env.OAUTH_KV.get(CONNECTION_MARKER_KEY);
		if (!raw) return null;
		return JSON.parse(raw) as ConnectionMarker;
	} catch {
		return null;
	}
}

export async function getValidAccessToken(env: Env): Promise<string> {
	try {
		return await makeStore(env).getValidAccessToken();
	} catch (e) {
		const internal = e instanceof Error ? e.message : String(e);

		// Microsoft rejected the refresh outright — consent revoked, secret
		// rotated in Azure AD, or the refresh token aged out. Reconnecting fixes
		// it, and the status is worth surfacing.
		const refreshFailed = /token refresh failed: (\d+)/.exec(internal);
		if (refreshFailed) {
			throw new Error(
				`Microsoft rejected the token refresh (HTTP ${refreshFailed[1]}). ` +
					"Visit /oauth/start to reconnect Outlook.",
			);
		}

		// Transient: another request held the refresh lock. Not an auth problem;
		// retrying usually succeeds.
		if (internal.includes("could not acquire lock")) {
			throw new Error(
				"Outlook token refresh is busy (another request holds the refresh lock). Retry shortly.",
			);
		}

		// No tokens. Distinguish "never connected" from "tokens stopped
		// decrypting and were discarded" — identical to the caller otherwise,
		// but they need completely different fixes.
		const marker = await readConnectionMarker(env);
		if (marker) {
			throw new Error(
				`Outlook was connected on ${marker.connected_at} (namespace ${marker.namespace}) ` +
					"but the stored tokens could no longer be decrypted and have been discarded. " +
					"This means MCP_APPROVAL_CODE changed, or the at-rest encryption changed underneath it. " +
					"Visit /oauth/start to reconnect Outlook.",
			);
		}

		throw new Error("Not authenticated. Visit /oauth/start to reconnect Outlook.");
	}
}

export async function clearTokens(env: Env): Promise<void> {
	await makeStore(env).clearTokens();
	await env.OAUTH_KV.delete(CONNECTION_MARKER_KEY).catch(() => {});
}
