import { createProxiedOAuthStore, type ProxiedOAuthStore } from "@bashco/mcp-toolkit";
import type { Env, OutlookTokenData } from "./types.js";

const TOKENS_KEY = "oauth_tokens";
const ENCRYPTION_NAMESPACE = "outlook-mcp-v2";

// Every Microsoft Graph delegated permission we ask for.
// `offline_access` is required for refresh tokens.
export const SCOPES = [
	"Mail.ReadWrite",
	"Mail.Send",
	"Calendars.ReadWrite",
	"Files.ReadWrite.All",
	"Contacts.ReadWrite",
	"MailboxSettings.ReadWrite",
	"Tasks.ReadWrite",
	"Sites.Read.All",
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
}

export async function getValidAccessToken(env: Env): Promise<string> {
	try {
		return await makeStore(env).getValidAccessToken();
	} catch {
		throw new Error("Not authenticated. Visit /oauth/start to reconnect Outlook.");
	}
}

export async function clearTokens(env: Env): Promise<void> {
	await makeStore(env).clearTokens();
}
