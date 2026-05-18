// Shared helpers for email composition: HTML sanitisation, attachment
// validation, and message building. Centralising these means every code path
// that sends or drafts mail goes through the same security checks.

import { ToolError } from "@bashco/mcp-toolkit";
import { graphRequestRaw } from "./graph.js";
import type { Env } from "./types.js";

// ── Limits ─────────────────────────────────────────────────────────────────────

export const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024; // 3 MB per file
export const MAX_TOTAL_ATTACHMENT_BYTES = 10 * 1024 * 1024; // 10 MB per email

// MIME allowlist. Anything not on this list is rejected. Keep tight.
const ALLOWED_MIME_TYPES = new Set([
	// Documents
	"application/pdf",
	"application/msword",
	"application/vnd.ms-excel",
	"application/vnd.ms-powerpoint",
	"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
	"application/vnd.openxmlformats-officedocument.presentationml.presentation",
	// Images
	// Note: image/svg+xml deliberately excluded. SVG is XML and can carry
	// <script> tags that execute when rendered; it is a known XSS vector.
	// Use PNG/JPG/WebP for inline images instead.
	"image/png",
	"image/jpeg",
	"image/jpg",
	"image/gif",
	"image/webp",
	// Text
	"text/plain",
	"text/csv",
	"text/html",
	"application/json",
	// Archives
	"application/zip",
]);

// ── Outbound HTML safety pass ──────────────────────────────────────────────────
// Strips script-execution vectors before mail leaves the account.
// This is defence-in-depth: it does not replace the recipient's mail-client
// sanitiser, but it ensures we never use this account to deliver weaponised
// HTML even if Claude is prompt-injected.
//
// Strengthened over the v1 sanitiser. Closes the bypasses found in audit:
//   - Polyglot tag closures (`<scr<script>ipt>`) — two-pass sanitisation.
//   - SVG attribute vectors (`<svg/onload=…>`) — whole <svg> block stripped.
//   - <style> blocks — stripped (CSS expression/url() XSS surface).
//   - HTML-entity-encoded colons in scheme (`javascript&#x3a;`) — normalised.
//   - CSS unicode escapes in scheme (`\6Aavascript:`) — normalised.
//   - `vbscript:`, `mhtml:`, `livescript:` URL schemes — blocked.
//   - `style="…"` attributes — stripped (background:url(...) etc.).
//   - `formaction="…"` on <button>/<input> — stripped.
//   - <math>/<foreignObject> — stripped.

function normaliseHtmlEntities(s: string): string {
	// Decode numeric (decimal + hex) and a small set of named entities so
	// scheme detection doesn't get fooled by `javascript&#x3a;...` etc.
	return s
		.replace(/&#x([0-9a-f]+);/gi, (_, hex) => {
			const cp = parseInt(hex, 16);
			return Number.isFinite(cp) && cp < 0x10000 ? String.fromCharCode(cp) : "";
		})
		.replace(/&#(\d+);/g, (_, dec) => {
			const cp = parseInt(dec, 10);
			return Number.isFinite(cp) && cp < 0x10000 ? String.fromCharCode(cp) : "";
		})
		.replace(/&colon;/gi, ":")
		.replace(/&#x?0*3a;?/gi, ":")
		.replace(/&Tab;|&NewLine;/g, "");
}

function normaliseCssEscapes(s: string): string {
	// CSS allows `\6A` etc. inside string/url values. Decode small range.
	return s.replace(/\\([0-9a-f]{1,6})\s?/gi, (_, hex) => {
		const cp = parseInt(hex, 16);
		return Number.isFinite(cp) && cp < 0x10000 ? String.fromCharCode(cp) : "";
	});
}

function singlePassSanitize(html: string): string {
	// Pre-normalise entities + CSS escapes so subsequent scheme-blocking regexes
	// see the real characters. We DON'T re-emit the normalised string — it's only
	// used as a basis for which tags/attributes to strip. But we DO emit a string
	// where scheme prefixes get blocked even if originally entity-encoded.
	const normalised = normaliseCssEscapes(normaliseHtmlEntities(html));

	// Block any URL-scheme prefix that appears in the *normalised* view — by
	// neutralising at the *raw* level if we can find it, or in the normalised
	// level if entity-encoded. We replace into the raw string so display HTML is
	// preserved aside from the dangerous bits.
	let out = html;
	if (/(?:javascript|vbscript|livescript|mhtml)\s*:/i.test(normalised)) {
		out = out.replace(/(?:javascript|vbscript|livescript|mhtml)\s*:/gi, "blocked:");
		// Also strip any entity-encoded variants that survived the raw-pass scan
		out = out.replace(/(?:&#x?[0-9a-f]+;)+/gi, (match) => {
			const decoded = normaliseHtmlEntities(match);
			if (/[a-z]:/i.test(decoded) || /^[a-z:]+$/i.test(decoded)) {
				return "";
			}
			return match;
		});
	}
	// CSS unicode-escape `javascript:` — kill at the raw level too
	out = out.replace(/\\6[Aa]\s?[aA]?[vV]?[aA]?[sS]?[cC]?[rR]?[iI]?[pP]?[tT]?:/g, "blocked:");

	return out
		// Remove entire <script>, <style>, <iframe>, <object>, <embed>, <form>,
		// <svg>, <math>, <foreignObject> blocks (including malformed/unclosed).
		.replace(/<script\b[\s\S]*?<\/script\s*>/gi, "")
		.replace(/<script\b[^>]*\/?>/gi, "") // self-closing or malformed open
		.replace(/<style\b[\s\S]*?<\/style\s*>/gi, "")
		.replace(/<style\b[^>]*\/?>/gi, "")
		.replace(/<iframe\b[\s\S]*?<\/iframe\s*>/gi, "")
		.replace(/<iframe\b[^>]*\/?>/gi, "")
		.replace(/<object\b[\s\S]*?<\/object\s*>/gi, "")
		.replace(/<object\b[^>]*\/?>/gi, "")
		.replace(/<embed\b[^>]*\/?>/gi, "")
		.replace(/<form\b[\s\S]*?<\/form\s*>/gi, "")
		.replace(/<form\b[^>]*\/?>/gi, "")
		.replace(/<svg\b[\s\S]*?<\/svg\s*>/gi, "")
		.replace(/<svg\b[^>]*\/?>/gi, "")
		.replace(/<math\b[\s\S]*?<\/math\s*>/gi, "")
		.replace(/<math\b[^>]*\/?>/gi, "")
		.replace(/<foreignObject\b[\s\S]*?<\/foreignObject\s*>/gi, "")
		.replace(/<link\b[^>]*\/?>/gi, "")
		.replace(/<meta\b[^>]*\/?>/gi, "")
		.replace(/<base\b[^>]*\/?>/gi, "")
		// Strip `on*=` event handler attributes (onclick, onload, onerror, etc.)
		// — handles quoted, single-quoted, unquoted, and slash-separator variants.
		.replace(/\s+on[a-z]+\s*=\s*"[^"]*"/gi, "")
		.replace(/\s+on[a-z]+\s*=\s*'[^']*'/gi, "")
		.replace(/\s+on[a-z]+\s*=\s*[^\s>]+/gi, "")
		.replace(/\/on[a-z]+\s*=\s*"[^"]*"/gi, "")
		.replace(/\/on[a-z]+\s*=\s*'[^']*'/gi, "")
		.replace(/\/on[a-z]+\s*=\s*[^\s>]+/gi, "")
		// Strip `style=`, `formaction=`, `action=` attributes — all are
		// attribute-only XSS sinks once script tags are gone.
		.replace(/\s+style\s*=\s*"[^"]*"/gi, "")
		.replace(/\s+style\s*=\s*'[^']*'/gi, "")
		.replace(/\s+style\s*=\s*[^\s>]+/gi, "")
		.replace(/\s+formaction\s*=\s*"[^"]*"/gi, "")
		.replace(/\s+formaction\s*=\s*'[^']*'/gi, "")
		.replace(/\s+formaction\s*=\s*[^\s>]+/gi, "")
		.replace(/\s+action\s*=\s*"[^"]*"/gi, "")
		.replace(/\s+action\s*=\s*'[^']*'/gi, "")
		.replace(/\s+action\s*=\s*[^\s>]+/gi, "")
		// Block any remaining `data:` URLs except the inline-image allowlist
		.replace(/\sdata\s*:(?!image\/(png|jpeg|jpg|gif|webp);base64,)/gi, " blocked:");
}

export function sanitizeOutboundHtml(html: string): string {
	// Two-pass — first pass may leave polyglot residue (e.g. `<scr<script>ipt>`
	// becomes `<script>` after the inner block is removed). Second pass catches it.
	return singlePassSanitize(singlePassSanitize(html));
}

// ── Attachment validation ──────────────────────────────────────────────────────

// An attachment can specify ONE of three sources:
//   - content_base64: raw base64 bytes inline (Claude has the bytes already)
//   - onedrive_path:  fetched server-side from the user's OneDrive
//   - url:            fetched server-side from a public HTTPS URL
// content_type is auto-detected for the latter two if omitted.
export interface AttachmentInput {
	name: string;
	content_base64?: string;
	onedrive_path?: string;
	url?: string;
	content_type?: string;
	is_inline?: boolean;
	content_id?: string;
}

interface GraphAttachment {
	"@odata.type": string;
	name: string;
	contentBytes: string;
	contentType: string;
	isInline?: boolean;
	contentId?: string;
}

export function validateAndBuildAttachments(attachments: unknown): GraphAttachment[] {
	if (!attachments) return [];
	if (!Array.isArray(attachments)) {
		throw ToolError.validation("attachments must be an array");
	}

	const built: GraphAttachment[] = [];
	let totalBytes = 0;

	for (const raw of attachments) {
		const a = raw as AttachmentInput;
		if (!a.name || !a.content_base64 || !a.content_type) {
			throw ToolError.validation(
				"Each attachment requires name, content_base64, and content_type",
			);
		}

		// MIME check
		const mime = a.content_type.toLowerCase();
		if (!ALLOWED_MIME_TYPES.has(mime)) {
			throw ToolError.validation(
				`MIME type not allowed: ${mime}. Allowed types: ${Array.from(ALLOWED_MIME_TYPES).join(", ")}`,
			);
		}

		// Filename extension defence-in-depth — block executable extensions
		// even if the MIME type lies.
		const dangerousExt = /\.(exe|bat|cmd|scr|msi|dll|ps1|vbs|com|cpl|jar|app)$/i;
		if (dangerousExt.test(a.name)) {
			throw ToolError.validation(`Filename extension not allowed: ${a.name}`);
		}

		// Size check (base64 expands ~33% — derive raw size)
		const rawBytes = Math.floor((a.content_base64.length * 3) / 4);
		if (rawBytes > MAX_ATTACHMENT_BYTES) {
			throw ToolError.validation(
				`Attachment "${a.name}" is ${(rawBytes / 1024 / 1024).toFixed(1)} MB. Max is ${MAX_ATTACHMENT_BYTES / 1024 / 1024} MB per file.`,
			);
		}
		totalBytes += rawBytes;
		if (totalBytes > MAX_TOTAL_ATTACHMENT_BYTES) {
			throw ToolError.validation(
				`Total attachments exceed ${MAX_TOTAL_ATTACHMENT_BYTES / 1024 / 1024} MB.`,
			);
		}

		// Inline image must have a content_id so the HTML body's cid: reference can resolve
		if (a.is_inline && !a.content_id) {
			throw ToolError.validation(
				`Inline attachment "${a.name}" requires a content_id (referenced in HTML as cid:THE_ID)`,
			);
		}

		built.push({
			"@odata.type": "#microsoft.graph.fileAttachment",
			name: a.name,
			contentBytes: a.content_base64,
			contentType: a.content_type,
			...(a.is_inline && a.content_id
				? { isInline: true, contentId: a.content_id }
				: {}),
		});
	}

	return built;
}

// ── Message body builder ───────────────────────────────────────────────────────

export function buildMessageBody(
	body: string,
	bodyType: string | undefined,
): { contentType: "Text" | "HTML"; content: string } {
	const isHtml = (bodyType ?? "text").toLowerCase() === "html";
	return {
		contentType: isHtml ? "HTML" : "Text",
		content: isHtml ? sanitizeOutboundHtml(body) : body,
	};
}

// ── Recipient list builder ─────────────────────────────────────────────────────

export function buildRecipients(
	csv: string | undefined,
): Array<{ emailAddress: { address: string } }> {
	if (!csv) return [];
	return csv
		.split(",")
		.map((e) => ({ emailAddress: { address: e.trim() } }))
		.filter((r) => r.emailAddress.address);
}

// ── Server-side attachment fetching (SSRF-safe) ────────────────────────────────

const URL_FETCH_TIMEOUT_MS = 10_000;

// Normalise IP-literal hostnames so a 32-bit decimal (`2130706433` = 127.0.0.1)
// or hex (`0x7f000001`) form can't bypass the dotted-decimal regex check.
// Also handles 4-segment dotted-decimal "as-is" and leaves DNS names untouched.
export function normaliseIpHostname(hostname: string): string {
	const lower = hostname.toLowerCase();

	// Hex: 0x7f000001
	const hexMatch = /^0x([0-9a-f]+)$/.exec(lower);
	if (hexMatch && hexMatch[1] !== undefined) {
		const n = parseInt(hexMatch[1], 16);
		if (Number.isFinite(n) && n >= 0 && n <= 0xffffffff) {
			return intToDottedDecimal(n);
		}
	}

	// Pure decimal: 2130706433
	if (/^\d+$/.test(lower)) {
		const n = parseInt(lower, 10);
		if (Number.isFinite(n) && n >= 0 && n <= 0xffffffff) {
			return intToDottedDecimal(n);
		}
	}

	return hostname;
}

function intToDottedDecimal(n: number): string {
	return [
		(n >>> 24) & 0xff,
		(n >>> 16) & 0xff,
		(n >>> 8) & 0xff,
		n & 0xff,
	].join(".");
}

// Block private/internal/loopback addresses to prevent SSRF.
// Note: this only catches URL-literal IPs, not DNS rebinding. Cloudflare
// Workers' fetch already refuses to connect to private ranges as additional
// defence, but we still validate up-front for clearer errors and to reject
// hostnames like "localhost".
export function rejectIfPrivateHost(hostname: string): void {
	const lower = hostname.toLowerCase();
	if (
		lower === "localhost" ||
		lower.endsWith(".localhost") ||
		lower.endsWith(".internal") ||
		lower.endsWith(".local")
	) {
		throw ToolError.validation(`Refusing to fetch from internal hostname: ${hostname}`);
	}

	// Normalise IP literals (handles decimal/hex 32-bit forms) before regex.
	const normalised = normaliseIpHostname(lower);

	const ipv4 = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/.exec(normalised);
	if (ipv4 && ipv4[1] !== undefined && ipv4[2] !== undefined) {
		const a = parseInt(ipv4[1], 10);
		const b = parseInt(ipv4[2], 10);
		const isPrivate =
			a === 10 || // 10.0.0.0/8
			a === 127 || // loopback
			a === 0 || // 0.0.0.0/8
			(a === 169 && b === 254) || // link-local / metadata
			(a === 172 && b >= 16 && b <= 31) || // 172.16.0.0/12
			(a === 192 && b === 168) || // 192.168.0.0/16
			a >= 224; // multicast / reserved
		if (isPrivate) throw ToolError.validation(`Refusing to fetch from private IP: ${hostname}`);
	}

	// IPv6 literal checks (Cloudflare URL might wrap in [])
	const stripped = lower.replace(/^\[|\]$/g, "");
	if (
		stripped === "::1" ||
		stripped.startsWith("fc") ||
		stripped.startsWith("fd") ||
		stripped.startsWith("fe80:") ||
		stripped.startsWith("::ffff:")
	) {
		throw ToolError.validation(`Refusing to fetch from private IPv6: ${hostname}`);
	}
}

export function validateExternalUrl(rawUrl: string): URL {
	let parsed: URL;
	try {
		parsed = new URL(rawUrl);
	} catch {
		throw ToolError.validation(`Invalid URL: ${rawUrl}`);
	}
	if (parsed.protocol !== "https:") {
		throw ToolError.validation(`Only https:// URLs are allowed (got ${parsed.protocol})`);
	}
	rejectIfPrivateHost(parsed.hostname);
	return parsed;
}

// Convert ArrayBuffer to base64. Chunked to avoid stack overflow on
// String.fromCharCode.apply with large arrays.
function arrayBufferToBase64(buf: ArrayBuffer): string {
	const bytes = new Uint8Array(buf);
	const chunkSize = 0x8000;
	let binary = "";
	for (let i = 0; i < bytes.length; i += chunkSize) {
		const chunk = bytes.subarray(i, i + chunkSize);
		binary += String.fromCharCode.apply(null, Array.from(chunk));
	}
	return btoa(binary);
}

// Read a Response body, enforcing a hard size cap mid-stream so a malicious
// or runaway server can't drain Worker memory.
async function readBodyWithCap(response: Response, maxBytes: number): Promise<ArrayBuffer> {
	const cl = response.headers.get("content-length");
	if (cl && parseInt(cl, 10) > maxBytes) {
		throw ToolError.validation(
			`Remote file is ${(parseInt(cl, 10) / 1024 / 1024).toFixed(1)} MB. Max is ${maxBytes / 1024 / 1024} MB.`,
		);
	}

	const reader = response.body?.getReader();
	if (!reader) {
		const buf = await response.arrayBuffer();
		if (buf.byteLength > maxBytes) {
			throw ToolError.validation(`Remote file exceeds ${maxBytes / 1024 / 1024} MB cap`);
		}
		return buf;
	}

	const chunks: Uint8Array[] = [];
	let total = 0;
	while (true) {
		const { done, value } = await reader.read();
		if (done) break;
		total += value.byteLength;
		if (total > maxBytes) {
			try {
				await reader.cancel();
			} catch {
				/* ignore */
			}
			throw ToolError.validation(`Remote file exceeds ${maxBytes / 1024 / 1024} MB cap`);
		}
		chunks.push(value);
	}

	const out = new Uint8Array(total);
	let offset = 0;
	for (const c of chunks) {
		out.set(c, offset);
		offset += c.byteLength;
	}
	return out.buffer;
}

// Fetch a public HTTPS URL with SSRF protection, timeout, and size cap.
// Redirects are handled manually so each hop's URL is re-validated against
// the SSRF blocklist — preventing a public URL that 302s to a private IP.
const MAX_REDIRECTS = 3;

async function fetchAsBase64FromUrl(
	rawUrl: string,
): Promise<{ content_base64: string; content_type: string }> {
	let currentUrl = rawUrl;
	validateExternalUrl(currentUrl);

	const controller = new AbortController();
	const timer = setTimeout(() => controller.abort(), URL_FETCH_TIMEOUT_MS);

	try {
		let response: Response | null = null;
		for (let hop = 0; hop <= MAX_REDIRECTS; hop++) {
			response = await fetch(currentUrl, {
				method: "GET",
				redirect: "manual",
				signal: controller.signal,
			});

			if ([301, 302, 303, 307, 308].includes(response.status)) {
				if (hop === MAX_REDIRECTS) {
					throw ToolError.validation(`Too many redirects (max ${MAX_REDIRECTS})`);
				}
				const location = response.headers.get("location");
				if (!location)
					throw ToolError.validation("Redirect response had no Location header");
				const next = new URL(location, currentUrl).toString();
				validateExternalUrl(next);
				currentUrl = next;
				continue;
			}

			break;
		}

		if (!response) throw new Error("No response received");
		if (!response.ok) {
			throw ToolError.validation(
				`URL fetch failed: ${response.status} ${response.statusText}`,
			);
		}
		const contentType = (
			response.headers.get("content-type") ?? "application/octet-stream"
		)
			.split(";")[0]!
			.trim();
		const buf = await readBodyWithCap(response, MAX_ATTACHMENT_BYTES);
		return { content_base64: arrayBufferToBase64(buf), content_type: contentType };
	} finally {
		clearTimeout(timer);
	}
}

// Fetch a file from the user's OneDrive by path, returning base64 + MIME.
async function fetchAsBase64FromOneDrive(
	env: Env,
	path: string,
): Promise<{ content_base64: string; content_type: string }> {
	const clean = path.replace(/^\/+/, "");
	const response = await graphRequestRaw(env, `/me/drive/root:/${encodeURI(clean)}:/content`);
	const contentType = (response.headers.get("content-type") ?? "application/octet-stream")
		.split(";")[0]!
		.trim();
	const buf = await readBodyWithCap(response, MAX_ATTACHMENT_BYTES);
	return { content_base64: arrayBufferToBase64(buf), content_type: contentType };
}

// ── Attachment resolver ────────────────────────────────────────────────────────
// Async because OneDrive/URL sources require fetches. Returns Graph-ready
// attachment objects (after security validation).

export async function resolveAndValidateAttachments(
	env: Env,
	raw: unknown,
): Promise<GraphAttachment[]> {
	if (!raw) return [];
	if (!Array.isArray(raw)) throw ToolError.validation("attachments must be an array");

	const resolved: AttachmentInput[] = [];
	for (const item of raw) {
		const a = item as AttachmentInput;
		if (!a.name) throw ToolError.validation('Each attachment requires a "name" field');

		const sources = [a.content_base64, a.onedrive_path, a.url].filter(Boolean).length;
		if (sources === 0) {
			throw ToolError.validation(
				`Attachment "${a.name}" needs one of: content_base64, onedrive_path, url`,
			);
		}
		if (sources > 1) {
			throw ToolError.validation(
				`Attachment "${a.name}" specifies multiple sources — pick one of content_base64, onedrive_path, url`,
			);
		}

		if (a.content_base64) {
			if (!a.content_type)
				throw ToolError.validation(
					`Attachment "${a.name}" with content_base64 requires content_type`,
				);
			resolved.push(a);
		} else if (a.onedrive_path) {
			const fetched = await fetchAsBase64FromOneDrive(env, a.onedrive_path);
			resolved.push({
				...a,
				content_base64: fetched.content_base64,
				content_type: a.content_type ?? fetched.content_type,
			});
		} else if (a.url) {
			const fetched = await fetchAsBase64FromUrl(a.url);
			resolved.push({
				...a,
				content_base64: fetched.content_base64,
				content_type: a.content_type ?? fetched.content_type,
			});
		}
	}

	return validateAndBuildAttachments(resolved);
}
