// Server-side email signature injection.
//
// The signature is appended by the Worker at send time so calling agents never
// reproduce it — they cannot paraphrase, truncate, or forget it.
//
// SECURITY / ORDERING NOTE (important, do not "simplify" this away):
// sanitizeOutboundHtml strips every `style=` attribute, because an attribute-only
// XSS sink is exactly what it exists to close. The signature is built entirely
// from inline styles, so running it through that sanitiser destroys it — logo
// sizing, the green divider, the CTA buttons, all of it, leaving a bare
// unstyled table. The signature must therefore be concatenated AFTER the
// caller's body has been sanitised.
//
// That is sound, not a bypass: the two strings have different trust levels.
// The body is agent-supplied and untrusted, so it is sanitised. The signature
// is operator-supplied deployment config (a Wrangler secret, set by whoever
// deploys the Worker) and is trusted by definition — an operator who can set
// secrets can already change the Worker's behaviour outright.

import { escapeHtml } from "@bashco/mcp-toolkit";
import { sanitizeOutboundHtml } from "./email-helpers.js";
import type { Env } from "./types.js";

/** Token inside SIGNATURE_HTML replaced with SIGNATURE_LOGO_URL at runtime. */
const LOGO_TOKEN = "__LOGO_URL__";

/**
 * Marker comment injected alongside the signature. Lets us recognise our own
 * previous injection exactly, without relying on text heuristics.
 */
const SIGNATURE_MARKER = "<!--outlook-mcp-signature-->";

/** Cap on the configured signature. Guards against a pathological secret. */
const MAX_SIGNATURE_BYTES = 64 * 1024;

export interface SignatureBuild {
	/** Final body content, ready to hand to Graph. Already sanitised. */
	content: string;
	/** Always HTML when a signature is present — see forceHtml note below. */
	contentType: "HTML" | "Text";
	/** Human-readable notes surfaced in the tool response. */
	notes: string[];
}

/**
 * Resolve the configured signature, substituting the logo URL.
 *
 * Returns null when SIGNATURE_HTML is unset or blank — an unconfigured
 * deployment sends unsigned mail rather than failing.
 */
export function resolveSignature(env: Env): string | null {
	const raw = env.SIGNATURE_HTML;
	if (typeof raw !== "string" || raw.trim() === "") return null;
	if (raw.length > MAX_SIGNATURE_BYTES) return null;

	const logo = typeof env.SIGNATURE_LOGO_URL === "string" ? env.SIGNATURE_LOGO_URL.trim() : "";

	if (!raw.includes(LOGO_TOKEN)) return raw;

	if (!logo) {
		// Logo URL missing: drop the <img> rather than emitting src="__LOGO_URL__",
		// which every mail client renders as a broken-image icon.
		return raw.replace(/<img\b[^>]*__LOGO_URL__[^>]*>/gi, "");
	}
	return raw.split(LOGO_TOKEN).join(escapeHtml(logo));
}

/**
 * Reduce HTML to comparable visible text: tags dropped, entities loosened,
 * whitespace collapsed, lowercased. Used only for the idempotency check.
 */
function visibleText(html: string): string {
	return html
		.replace(/<[^>]*>/g, " ")
		.replace(/&nbsp;/gi, " ")
		.replace(/&amp;/gi, "&")
		.replace(/\s+/g, " ")
		.trim()
		.toLowerCase();
}

/**
 * Minimum length for a signature phrase to count as identifying. Short enough
 * to catch a tagline, long enough that ordinary prose will not collide with it.
 */
const MIN_PHRASE_LENGTH = 15;

/**
 * Distinctive phrases derived from the CONFIGURED signature, so nothing about
 * any particular signature is hardcoded here. The visible text is split on
 * sentence and separator boundaries; for a typical block these come out as the
 * name line, the tagline, and the CTA labels.
 */
function signaturePhrases(signature: string): string[] {
	return visibleText(signature)
		.split(/\s*[.|•·]\s*|\s{2,}/)
		.map((s) => s.trim())
		.filter((s) => s.length >= MIN_PHRASE_LENGTH);
}

/**
 * True when the body already carries this signature — either our own marker
 * (we injected it on a previous pass) or one of the signature's distinctive
 * phrases (an agent pasted a copy manually and also set the flag).
 *
 * Matching on any single phrase rather than the whole block is deliberate: an
 * agent that re-typed the signature with different markup, or included only the
 * tagline, still counts as already-signed. A false positive costs one missing
 * signature plus a note in the response; a false negative sends the recipient a
 * visibly doubled block.
 */
export function bodyContainsSignature(body: string, signature: string): boolean {
	if (body.includes(SIGNATURE_MARKER)) return true;
	const phrases = signaturePhrases(signature);
	if (phrases.length === 0) return false;
	const text = visibleText(body);
	return phrases.some((p) => text.includes(p));
}

/**
 * Outlook marks where new content ends and the quoted original begins.
 * `appendonsend` is Microsoft's own canonical marker; `divRplyFwdMsg` is the
 * older/Outlook-desktop form. Checked in that order.
 */
const QUOTE_BOUNDARIES = [
	/<div\b[^>]*\bid\s*=\s*["']?appendonsend["']?[^>]*>/i,
	/<div\b[^>]*\bid\s*=\s*["']?divRplyFwdMsg["']?[^>]*>/i,
	/<hr\b[^>]*\bid\s*=\s*["']?stopSpelling["']?[^>]*>/i,
];

/**
 * Index at which the quoted original begins, or null when this body carries no
 * recognisable quote block.
 */
export function findQuoteBoundary(html: string): number | null {
	for (const boundary of QUOTE_BOUNDARIES) {
		const match = boundary.exec(html);
		if (match) return match.index;
	}
	return null;
}

/**
 * Insert HTML into a Graph-generated reply/forward body ABOVE the quoted
 * original, keeping the quote intact below it.
 *
 * Falls back to appending when no quote boundary is recognised — better a
 * signature at the bottom than content silently dropped.
 */
export function insertBeforeQuote(existingHtml: string, insert: string): string {
	const at = findQuoteBoundary(existingHtml);
	if (at !== null) {
		return `${existingHtml.slice(0, at)}${insert}${existingHtml.slice(at)}`;
	}
	return existingHtml.trim() === "" ? insert : `${existingHtml}<br><br>${insert}`;
}

/** Insert the signature above the quoted original. */
export function insertAboveQuote(existingHtml: string, signature: string): string {
	const block = `${SIGNATURE_MARKER}${signature}`;
	const at = findQuoteBoundary(existingHtml);
	if (at !== null) {
		return `${existingHtml.slice(0, at)}<br><br>${block}${existingHtml.slice(at)}`;
	}
	return existingHtml.trim() === "" ? block : `${existingHtml}<br><br>${block}`;
}

/**
 * Render an author-supplied body to HTML, applying the same trust rules as
 * buildBodyWithSignature: agent HTML is sanitised, plain text is escaped and
 * newline-converted. Exported so reply/forward composition can reuse it.
 */
export function renderBodyHtml(body: string, bodyType: string | undefined): string {
	return (bodyType ?? "text").toLowerCase() === "html"
		? sanitizeOutboundHtml(body)
		: textToHtml(body);
}

/** The marker + signature block, ready to concatenate. */
export function signatureBlock(signature: string): string {
	return `${SIGNATURE_MARKER}${signature}`;
}

/**
 * Build a message body with the signature appended.
 *
 * Forces HTML: a signature inside a plain-text body renders as visible raw
 * markup, so the combination is never allowed. When the caller explicitly asked
 * for `body_type: "text"` the override is recorded in `notes` and surfaced in
 * the tool response rather than being applied silently.
 */
export function buildBodyWithSignature(
	body: string | undefined,
	bodyType: string | undefined,
	signature: string,
): SignatureBuild {
	const notes: string[] = [];
	const raw = body ?? "";
	const wasHtml = (bodyType ?? "text").toLowerCase() === "html";

	if (bodyType !== undefined && !wasHtml) {
		notes.push(
			"body_type was overridden from 'text' to 'html' because include_signature was set — " +
				"a signature in a plain-text body would render as raw markup.",
		);
	}

	// Emptiness is judged on the RAW body, before newline conversion. A
	// whitespace-only body like "\n\t " would otherwise become "<br>\t" and
	// stop looking empty, producing a stray leading break above the signature.
	const isEmpty = raw.trim() === "";

	// Already signed: return the body unchanged rather than stacking a copy.
	if (!isEmpty && bodyContainsSignature(raw, signature)) {
		notes.push("Signature already present in the supplied body — not appended again.");
		return {
			content: wasHtml ? sanitizeOutboundHtml(raw) : textToHtml(raw),
			contentType: "HTML",
			notes,
		};
	}

	const block = `${SIGNATURE_MARKER}${signature}`;
	if (isEmpty) return { content: block, contentType: "HTML", notes };

	// Sanitise the untrusted body FIRST, then concatenate the trusted signature.
	const safeBody = wasHtml ? sanitizeOutboundHtml(raw) : textToHtml(raw);
	return { content: `${safeBody}<br><br>${block}`, contentType: "HTML", notes };
}

/**
 * Escape a plain-text body and convert newlines to <br>.
 *
 * Escaping first matters: forcing HTML on text that contains `<` or `&` would
 * otherwise let it be reinterpreted as markup. Newlines then become <br> so
 * the author's line breaks survive the forced HTML switch.
 */
function textToHtml(text: string): string {
	return escapeHtml(text).replace(/\r\n/g, "\n").replace(/\n/g, "<br>");
}

/**
 * Convenience wrapper: returns null when signing is not requested or not
 * configured, so callers keep their existing code path unchanged.
 */
export function maybeSign(
	env: Env,
	include: boolean | undefined,
	body: string | undefined,
	bodyType: string | undefined,
): SignatureBuild | null {
	if (!include) return null;
	const signature = resolveSignature(env);
	if (!signature) return null;
	return buildBodyWithSignature(body, bodyType, signature);
}
