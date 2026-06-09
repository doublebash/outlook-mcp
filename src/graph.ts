import { ToolError, createUpstreamClient } from "@bashco/mcp-toolkit";
import { getValidAccessToken } from "./auth.js";
import type { Env } from "./types.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

interface GraphErrorEnvelope {
	error?: {
		code?: string;
		message?: string;
		innerError?: { "request-id"?: string; date?: string };
	};
}

// Microsoft Graph error parser. Extracts `error.code` (the stable
// machine-readable identifier like `ErrorItemNotFound`, `ErrorAccessDenied`,
// `ErrorMessageSizeExceeded`) and surfaces it to Claude alongside the message,
// so Claude can branch on the failure type.
function parseGraphError(internal: string): { code?: string; message?: string } | null {
	// internal is shaped like `Graph 404: {...json...}`
	const m = internal.match(/Graph \d+: ([\s\S]+)$/);
	if (!m || !m[1]) return null;
	try {
		const parsed = JSON.parse(m[1]) as GraphErrorEnvelope;
		if (!parsed.error) return null;
		const result: { code?: string; message?: string } = {};
		if (parsed.error.code) result.code = parsed.error.code;
		if (parsed.error.message) result.message = parsed.error.message;
		return result;
	} catch {
		return null;
	}
}

type GraphInit = {
	method: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
	path: string;
	body?: unknown;
	query?: Record<string, string | number | undefined>;
	headers?: Record<string, string>;
};

async function graphFetch(env: Env, init: GraphInit): Promise<unknown> {
	const token = await getValidAccessToken(env);
	const client = createUpstreamClient({
		upstreamName: "Graph",
		baseUrl: GRAPH_BASE,
		buildHeaders: async () => ({
			Authorization: `Bearer ${token}`,
			Accept: "application/json",
		}),
	});

	try {
		return await client.fetch({
			method: init.method,
			path: init.path,
			...(init.query !== undefined ? { query: init.query } : {}),
			...(init.body !== undefined ? { body: init.body } : {}),
			...(init.headers !== undefined ? { headers: init.headers } : {}),
		});
	} catch (e) {
		// Reshape Graph's odata error envelope into a friendlier userMessage.
		if (e instanceof ToolError && typeof e.internalMessage === "string") {
			const parsed = parseGraphError(e.internalMessage);
			if (parsed?.code || parsed?.message) {
				const detail = parsed.code
					? `${parsed.code}${parsed.message ? `: ${parsed.message}` : ""}`
					: parsed.message ?? "";
				throw new ToolError({
					userMessage: `Outlook error ${e.status ?? "unknown"}: ${detail}`,
					internalMessage: e.internalMessage,
					...(e.status !== undefined ? { status: e.status } : {}),
					upstreamName: "Graph",
				});
			}
		}
		throw e;
	}
}

export async function graphGet(
	env: Env,
	path: string,
	query?: Record<string, string | number | undefined>,
	headers?: Record<string, string>,
): Promise<unknown> {
	return graphFetch(env, {
		method: "GET",
		path,
		...(query !== undefined ? { query } : {}),
		...(headers !== undefined ? { headers } : {}),
	});
}

export async function graphPost(env: Env, path: string, body?: unknown): Promise<unknown> {
	return graphFetch(env, {
		method: "POST",
		path,
		...(body !== undefined ? { body } : {}),
	});
}

export async function graphPatch(env: Env, path: string, body?: unknown): Promise<unknown> {
	return graphFetch(env, {
		method: "PATCH",
		path,
		...(body !== undefined ? { body } : {}),
	});
}

export async function graphDelete(env: Env, path: string): Promise<unknown> {
	return graphFetch(env, { method: "DELETE", path });
}

// ── Raw-response variant ──────────────────────────────────────────────────────
// Used when we need the binary body (OneDrive attachment downloads). The
// toolkit's createUpstreamClient drops non-JSON responses on the floor, so we
// fetch directly here. Caller is responsible for consuming the body.

export async function graphRequestRaw(env: Env, path: string): Promise<Response> {
	const token = await getValidAccessToken(env);
	const url = `${GRAPH_BASE}${path.startsWith("/") ? path : `/${path}`}`;
	const response = await fetch(url, {
		method: "GET",
		headers: { Authorization: `Bearer ${token}` },
		redirect: "follow",
	});

	if (!response.ok) {
		let errMsg = response.statusText;
		let errCode: string | undefined;
		try {
			const data = (await response.clone().json()) as GraphErrorEnvelope;
			errMsg = data.error?.message ?? errMsg;
			errCode = data.error?.code;
		} catch {
			// not JSON
		}
		throw new ToolError({
			userMessage: `Outlook error ${response.status}: ${errCode ?? errMsg}`,
			internalMessage: `Graph ${response.status}: ${errCode ?? ""} ${errMsg}`,
			status: response.status,
			upstreamName: "Graph",
		});
	}

	return response;
}

// PUT a binary body to a Graph endpoint with an explicit Content-Type — used
// by OneDrive simple upload (`/me/drive/root:/<path>:/content`). The toolkit
// client serialises body as JSON, which is wrong for opaque file bytes.
// Parses the success-case driveItem JSON so callers can return metadata.
export async function graphPutBinary(
	env: Env,
	path: string,
	body: ArrayBuffer | Uint8Array,
	contentType: string,
): Promise<unknown> {
	const token = await getValidAccessToken(env);
	const url = `${GRAPH_BASE}${path.startsWith("/") ? path : `/${path}`}`;
	const response = await fetch(url, {
		method: "PUT",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": contentType,
		},
		body: body as BodyInit,
		redirect: "follow",
	});

	if (!response.ok) {
		let errMsg = response.statusText;
		let errCode: string | undefined;
		try {
			const data = (await response.clone().json()) as GraphErrorEnvelope;
			errMsg = data.error?.message ?? errMsg;
			errCode = data.error?.code;
		} catch {
			// not JSON
		}
		throw new ToolError({
			userMessage: `Outlook error ${response.status}: ${errCode ?? errMsg}`,
			internalMessage: `Graph ${response.status}: ${errCode ?? ""} ${errMsg}`,
			status: response.status,
			upstreamName: "Graph",
		});
	}

	return response.json();
}
