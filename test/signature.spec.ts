import { describe, it, expect } from "vitest";
import {
	buildBodyWithSignature,
	bodyContainsSignature,
	insertAboveQuote,
	maybeSign,
	resolveSignature,
} from "../src/signature.js";

// A miniature stand-in for the real signature block: inline styles, a logo
// token, a tagline, and a CTA link. Deliberately NOT the operator's real
// signature — nothing about it is hardcoded in src/.
const SIG_TEMPLATE = `<table style="border-collapse:collapse;"><tr><td><img src="__LOGO_URL__" width="70" style="display:block;" alt="Logo"></td><td style="border-left:2px solid #44B386;"><div style="font-weight:700;">Test Sender</div><div style="letter-spacing:0.4px;">WIN THE LISTING. SELL THE PROPERTY.</div><a href="https://example.com/quote" style="background:#44B386;">Get an instant quote</a></td></tr></table>`;

const LOGO = "https://cdn.example.com/logo.png";

function env(overrides: Record<string, unknown> = {}) {
	return {
		SIGNATURE_HTML: SIG_TEMPLATE,
		SIGNATURE_LOGO_URL: LOGO,
		...overrides,
	} as never;
}

const SIG = SIG_TEMPLATE.split("__LOGO_URL__").join(LOGO);

describe("resolveSignature — configuration", () => {
	it("substitutes __LOGO_URL__ with the configured logo", () => {
		const out = resolveSignature(env())!;
		expect(out).toContain(`src="${LOGO}"`);
		expect(out).not.toContain("__LOGO_URL__");
	});

	it("returns null when SIGNATURE_HTML is unset — unsigned, never an error", () => {
		expect(resolveSignature(env({ SIGNATURE_HTML: undefined }))).toBeNull();
		expect(resolveSignature(env({ SIGNATURE_HTML: "   " }))).toBeNull();
	});

	it("drops the <img> rather than emitting a literal token when the logo URL is missing", () => {
		const out = resolveSignature(env({ SIGNATURE_LOGO_URL: undefined }))!;
		// A literal src="__LOGO_URL__" renders as a broken-image icon in every client.
		expect(out).not.toContain("__LOGO_URL__");
		expect(out).not.toContain("<img");
		expect(out).toContain("Test Sender");
	});

	it("escapes the logo URL so it cannot break out of the src attribute", () => {
		const out = resolveSignature(env({ SIGNATURE_LOGO_URL: 'https://x/a.png" onerror="alert(1)' }))!;
		expect(out).not.toContain('onerror="alert(1)"');
	});
});

describe("buildBodyWithSignature — core behaviour", () => {
	it("appends the signature separated by <br><br>", () => {
		const out = buildBodyWithSignature("Hi there", "text", SIG);
		expect(out.content).toContain("Hi there<br><br>");
		expect(out.content).toContain("Test Sender");
	});

	it("always forces HTML", () => {
		expect(buildBodyWithSignature("Hi", "text", SIG).contentType).toBe("HTML");
		expect(buildBodyWithSignature("Hi", undefined, SIG).contentType).toBe("HTML");
		expect(buildBodyWithSignature("<p>Hi</p>", "html", SIG).contentType).toBe("HTML");
	});

	it("reports the body_type override instead of applying it silently", () => {
		const out = buildBodyWithSignature("Hi", "text", SIG);
		expect(out.notes.join(" ")).toMatch(/overridden from 'text' to 'html'/);
	});

	it("does not add a note when the caller already asked for html", () => {
		expect(buildBodyWithSignature("<p>Hi</p>", "html", SIG).notes).toHaveLength(0);
	});
});

// This is the load-bearing assertion for the whole design. sanitizeOutboundHtml
// strips every style= attribute, so if the signature were routed through it the
// block would arrive unstyled — no logo sizing, no divider, no buttons.
describe("buildBodyWithSignature — sanitisation boundary", () => {
	it("preserves the signature's inline styles", () => {
		const out = buildBodyWithSignature("Hello", "text", SIG);
		expect(out.content).toContain('style="border-collapse:collapse;"');
		expect(out.content).toContain("border-left:2px solid #44B386");
		expect(out.content).toContain("background:#44B386");
	});

	it("still sanitises the caller-supplied HTML body", () => {
		const out = buildBodyWithSignature(
			'<p onclick="steal()">Hi</p><script>alert(1)</script>',
			"html",
			SIG,
		);
		expect(out.content).not.toContain("<script");
		expect(out.content).not.toContain("onclick");
		// ...while the trusted signature came through intact.
		expect(out.content).toContain("border-left:2px solid #44B386");
	});

	it("strips style attributes from the untrusted body but not the signature", () => {
		const out = buildBodyWithSignature('<p style="color:red">Hi</p>', "html", SIG);
		expect(out.content).not.toContain("color:red");
		expect(out.content).toContain("border-collapse:collapse");
	});
});

describe("buildBodyWithSignature — edge cases", () => {
	it("sends the signature alone with no leading spacer for an empty body", () => {
		for (const body of ["", "   ", "\n\t ", undefined]) {
			const out = buildBodyWithSignature(body, "text", SIG);
			expect(out.content.startsWith("<br>")).toBe(false);
			expect(out.content).not.toContain("<br><br>");
			expect(out.content).toContain("Test Sender");
		}
	});

	it("converts newlines in a plain-text body to <br>", () => {
		const out = buildBodyWithSignature("line one\nline two\nline three", "text", SIG);
		expect(out.content).toContain("line one<br>line two<br>line three");
	});

	it("normalises CRLF as a single <br>", () => {
		const out = buildBodyWithSignature("a\r\nb", "text", SIG);
		expect(out.content).toContain("a<br>b");
		expect(out.content).not.toContain("<br><br>b");
	});

	it("escapes markup in a plain-text body before forcing HTML", () => {
		// Forcing html on text without escaping would let "<b>" become real markup.
		const out = buildBodyWithSignature("5 < 6 & <b>not bold</b>", "text", SIG);
		expect(out.content).not.toContain("<b>not bold</b>");
		expect(out.content).toContain("&lt;b&gt;");
	});
});

describe("buildBodyWithSignature — idempotency", () => {
	it("does not append twice when the agent already pasted the signature", () => {
		const already = `Hi there<br><br>${SIG}`;
		const out = buildBodyWithSignature(already, "html", SIG);
		const occurrences = out.content.split("WIN THE LISTING").length - 1;
		expect(occurrences).toBe(1);
		expect(out.notes.join(" ")).toMatch(/already present/);
	});

	it("recognises a paraphrased-but-present signature by its text fingerprint", () => {
		// Agent re-typed the block with different markup but the same tagline.
		const already = "<div>Regards</div><div>WIN THE LISTING. SELL THE PROPERTY.</div>";
		expect(bodyContainsSignature(already, SIG)).toBe(true);
	});

	it("recognises our own previous injection via the marker comment", () => {
		const once = buildBodyWithSignature("Hi", "text", SIG).content;
		const twice = buildBodyWithSignature(once, "html", SIG);
		expect(twice.content.split("WIN THE LISTING").length - 1).toBe(1);
	});

	it("does not false-positive on an unrelated body", () => {
		expect(bodyContainsSignature("Thanks, talk soon.", SIG)).toBe(false);
	});
});

describe("insertAboveQuote — reply/forward placement", () => {
	const QUOTED = '<div id="divRplyFwdMsg">From: someone@example.com<br>Sent: ...</div>';

	it("inserts above Outlook's appendonsend boundary", () => {
		const body = `<div>Draft area</div><div id="appendonsend"></div>${QUOTED}`;
		const out = insertAboveQuote(body, SIG);
		expect(out.indexOf("Test Sender")).toBeLessThan(out.indexOf("appendonsend"));
	});

	it("inserts above the divRplyFwdMsg boundary when appendonsend is absent", () => {
		const out = insertAboveQuote(`<div>Draft area</div>${QUOTED}`, SIG);
		expect(out.indexOf("Test Sender")).toBeLessThan(out.indexOf("divRplyFwdMsg"));
	});

	it("keeps the quoted original intact", () => {
		const out = insertAboveQuote(`<div>Hi</div>${QUOTED}`, SIG);
		expect(out).toContain("From: someone@example.com");
	});

	it("falls back to appending when no quote boundary is recognised", () => {
		const out = insertAboveQuote("<div>Just a body</div>", SIG);
		expect(out.indexOf("Just a body")).toBeLessThan(out.indexOf("Test Sender"));
	});

	it("emits the signature alone for an empty draft body", () => {
		expect(insertAboveQuote("", SIG).startsWith("<br><br>")).toBe(false);
	});
});

describe("maybeSign — opt-in gating", () => {
	it("returns null when include_signature is absent or false", () => {
		expect(maybeSign(env(), undefined, "Hi", "text")).toBeNull();
		expect(maybeSign(env(), false, "Hi", "text")).toBeNull();
	});

	it("returns null (not an error) when the deployment has no signature configured", () => {
		expect(maybeSign(env({ SIGNATURE_HTML: undefined }), true, "Hi", "text")).toBeNull();
	});

	it("signs when requested and configured", () => {
		const out = maybeSign(env(), true, "Hi", "text")!;
		expect(out.contentType).toBe("HTML");
		expect(out.content).toContain("Test Sender");
	});
});
