import { z } from "zod";

// ── Shared tool descriptions ──────────────────────────────────────────────────

export const BODY_TYPE_DESC =
	'Body format: "text" (default) or "html". When "html", basic rich-text tags are allowed (b, i, u, p, br, ul, ol, li, a, img, table, etc.). Script, style, event-handler attributes, and dangerous URL schemes are stripped before sending.';

export const ATTACHMENTS_DESC =
	'Attachments array. Each item must have a name plus EXACTLY ONE source: (a) content_base64 + content_type for inline base64 bytes, OR (b) onedrive_path (e.g. "Documents/report.pdf") to fetch from the user\'s OneDrive server-side, OR (c) url (https only) to fetch from a public web URL server-side. Optional: is_inline + content_id for inline images (reference them in HTML body as cid:THE_ID). Max 3 MB per file, 10 MB total. Allowed types: pdf, docx, xlsx, pptx, png, jpg, gif, webp, txt, csv, html, json, zip. PREFER onedrive_path or url over content_base64 — the bytes never travel through the conversation, which is faster and cheaper.';

// ── Shared Zod schemas ────────────────────────────────────────────────────────

// Each attachment supplies one of three sources. We validate the *union* at
// schema time (any of the three keys present), then leave the per-attachment
// "exactly one source" enforcement to resolveAndValidateAttachments — which is
// where the runtime checks for size, MIME allowlist, and extension blocklist live.
export const attachmentSchema = z.object({
	name: z.string().min(1).max(256),
	content_base64: z.string().min(1).optional(),
	onedrive_path: z.string().min(1).max(1024).optional(),
	url: z.string().url().max(2048).optional(),
	content_type: z.string().min(1).max(128).optional(),
	is_inline: z.boolean().optional(),
	content_id: z.string().min(1).max(128).optional(),
});

export const recurrenceSchema = z.object({
	pattern: z.enum(["daily", "weekly", "monthly", "yearly"]),
	interval: z.number().int().min(1).max(366).optional(),
	days_of_week: z
		.array(
			z.enum([
				"sunday",
				"monday",
				"tuesday",
				"wednesday",
				"thursday",
				"friday",
				"saturday",
			]),
		)
		.min(1)
		.max(7)
		.optional(),
	day_of_month: z.number().int().min(1).max(31).optional(),
	month: z.number().int().min(1).max(12).optional(),
	end_date: z
		.string()
		.regex(/^\d{4}-\d{2}-\d{2}$/, "expected YYYY-MM-DD")
		.optional(),
	occurrences: z.number().int().min(1).max(9999).optional(),
});

// ── Audit logging ─────────────────────────────────────────────────────────────
// Structured JSON line emitted to Cloudflare's observability (3-day retention).
// Format is deliberately kept identical to v1 — top-level `audit` key with the
// event name as value — so existing dashboard filters (`audit:email_sent` etc.)
// continue to work. Nested fields (e.g. attachment_sources) are preserved as
// objects, which the toolkit's structured logger (scalar-only) doesn't support.

export function auditLog(event: string, details: Record<string, unknown>): void {
	console.log(
		JSON.stringify({ audit: event, ts: new Date().toISOString(), ...details }),
	);
}

// Summarise attachment sources without leaking full paths/URLs.
// Returns e.g. {onedrive: 2, url: 1, base64: 0}.
export function summariseAttachmentSources(raw: unknown): Record<string, number> {
	if (!Array.isArray(raw)) return {};
	const out: Record<string, number> = { onedrive: 0, url: 0, base64: 0 };
	for (const item of raw) {
		const a = item as Record<string, unknown>;
		if (a.onedrive_path) out.onedrive!++;
		else if (a.url) out.url!++;
		else if (a.content_base64) out.base64!++;
	}
	return out;
}

// ── OData filter helpers ──────────────────────────────────────────────────────

// Escape single quotes for OData string literals (e.g. `O'Brien` → `O''Brien`)
// then URI-encode if the value will sit inside a URL filter parameter.
export function escapeOdataString(s: string): string {
	return s.replace(/'/g, "''");
}

// Encode a folder/file path for the OneDrive `/me/drive/root:/<path>:/...`
// template. Each segment is URI-encoded individually so slashes stay
// path-separators and special chars don't bleed through.
export function encodeOneDrivePath(path: string): string {
	return path
		.split("/")
		.filter((seg) => seg.length > 0)
		.map((seg) => encodeURIComponent(seg))
		.join("/");
}
