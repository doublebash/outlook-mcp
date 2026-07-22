import type { CloudflareRateLimiter } from "@bashco/mcp-toolkit";

export interface Env {
	OAUTH_KV: KVNamespace;
	MICROSOFT_CLIENT_ID: string;
	MICROSOFT_CLIENT_SECRET: string;
	MICROSOFT_TENANT_ID: string;
	WORKER_URL: string;
	/**
	 * Operator-only approval code. Gates `/approve` and `/oauth/start`.
	 * Also the HKDF-derived encryption secret for Microsoft tokens stored in KV.
	 * NEVER returned to clients. NEVER used as a runtime bearer.
	 */
	MCP_APPROVAL_CODE: string;

	/**
	 * Optional email signature block, appended by the Worker when a send tool is
	 * called with `include_signature: true`. Set with `wrangler secret put` — it
	 * is deployment config, never a committed source file, so a fork does not
	 * send mail carrying someone else's name and booking links.
	 *
	 * May contain the token `__LOGO_URL__`, substituted with SIGNATURE_LOGO_URL.
	 * Unset means send unsigned mail; it is never an error.
	 */
	SIGNATURE_HTML?: string;
	/** Publicly reachable HTTPS URL of the signature logo. */
	SIGNATURE_LOGO_URL?: string;

	RATE_LIMIT_APPROVE: CloudflareRateLimiter;
	RATE_LIMIT_TOKEN: CloudflareRateLimiter;
	RATE_LIMIT_REGISTER: CloudflareRateLimiter;
	RATE_LIMIT_MCP: CloudflareRateLimiter;
	RATE_LIMIT_OAUTH_START: CloudflareRateLimiter;
}

/** Microsoft token shape stored encrypted at rest. */
export interface OutlookTokenData {
	access_token: string;
	refresh_token: string;
	expires_at: number; // Unix ms
}
