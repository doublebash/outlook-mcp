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

// ── Query strings ─────────────────────────────────────────────────────────────
// The toolkit's upstream client builds query strings with URLSearchParams, which
// percent-encodes `$` to `%24` — it only leaves ASCII alphanumerics and `*-._`
// alone. Graph's Outlook/Exchange backends (/me/messages, /me/events,
// /me/contacts, /me/drive) decode that and carry on, so `%24select` went
// unnoticed for as long as those were the only endpoints we called.
//
// Microsoft To Do does not. `/me/todo/*` is fronted by a request broker that
// never decodes the parameter *name*, so it fails to parse the URI and answers
// `400 invalidRequest` with `innerError.code` of `RequestBroker--ParseUri` — a
// generic error that names neither the parameter nor the property. Any To Do
// call carrying a query string died on it; the ones without a query string
// (resolving the default list, POSTing a new task) were unaffected.
//
// So the query string is built here, with OData option names left literal the
// way Microsoft's own documentation writes them. Values are still percent-
// encoded, minus the two characters OData syntax needs to stay readable and
// which are legal unencoded in a query string per RFC 3986: `,` separating
// $select fields and $filter function arguments, and `/` in nested property
// paths like `start/dateTime`.
function encodeGraphValue(value: string): string {
	return encodeURIComponent(value).replace(/%2C/g, ",").replace(/%2F/g, "/");
}

export function buildGraphPath(
	path: string,
	query?: Record<string, string | number | undefined>,
): string {
	if (!query) return path;
	const parts: string[] = [];
	for (const [key, value] of Object.entries(query)) {
		if (value === undefined || value === null || value === "") continue;
		parts.push(`${key}=${encodeGraphValue(String(value))}`);
	}
	return parts.length > 0 ? `${path}?${parts.join("&")}` : path;
}

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
			// Query folded into the path so it keeps this module's encoding — handing
			// `query` to the client would put it back through URLSearchParams.
			method: init.method,
			path: buildGraphPath(init.path, init.query),
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

// ── Paging ────────────────────────────────────────────────────────────────────
// Graph advertises the next page as an absolute `@odata.nextLink`. Microsoft's
// paging guidance is explicit that the whole URL has to be reused as-is:
//
//   "Depending on the API that the query is being performed against, the
//    @odata.nextLink URL value contains either a $skiptoken or a $skip query
//    parameter. Any other query parameters that were present in the original
//    request are also encoded in this URL. Don't try to extract the $skiptoken
//    or $skip value and use it in a different request."
//
// Pulling one named token out of the link and rebuilding the query looks like it
// works — until it meets an endpoint that pages with the other form, where the
// lookup returns null and paging silently stops after page one. /me/events is
// one such endpoint. Silent truncation is worse than an error, so we follow the
// link instead of second-guessing it.
//
// The link still gets validated before it's followed: it arrives inside a
// response body, and a bearer token must never be sent to a host we didn't mean
// to call.
export function resolveNextLinkPath(nextLink: string): string {
	let url: URL;
	try {
		url = new URL(nextLink);
	} catch {
		throw new ToolError({
			userMessage: "Outlook returned a malformed pagination link.",
			internalMessage: `Unparseable @odata.nextLink: ${nextLink}`,
			upstreamName: "Graph",
		});
	}

	const base = new URL(GRAPH_BASE);
	const withinBase =
		url.origin === base.origin &&
		(url.pathname === base.pathname || url.pathname.startsWith(`${base.pathname}/`));
	if (!withinBase) {
		throw new ToolError({
			userMessage: "Outlook returned a pagination link pointing somewhere unexpected.",
			internalMessage: `@odata.nextLink outside ${GRAPH_BASE}: ${url.origin}${url.pathname}`,
			upstreamName: "Graph",
		});
	}

	// The upstream client concatenates baseUrl + path, so hand back everything
	// after the base — query string included, untouched.
	return `${url.pathname.slice(base.pathname.length)}${url.search}`;
}

export async function graphGetNextLink(env: Env, nextLink: string): Promise<unknown> {
	return graphFetch(env, { method: "GET", path: resolveNextLinkPath(nextLink) });
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
