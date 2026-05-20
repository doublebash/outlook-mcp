import { describe, it, expect } from "vitest";
import { SELF } from "cloudflare:test";
import { htmlToText } from "../src/sanitize.js";
import {
	sanitizeOutboundHtml,
	rejectIfPrivateHost,
	normaliseIpHostname,
	validateAndBuildAttachments,
} from "../src/email-helpers.js";
import { buildRecurrence } from "../src/calendar-helpers.js";
import { parseVtt } from "../src/transcript-helpers.js";
import { extractJoinUrlFromBody } from "../src/tools/meetings.js";
import {
	normaliseStartDate,
	normaliseEndDate,
} from "../src/tools/calendar.js";
import { computeSelfReplyRecipients } from "../src/tools/email.js";

// ── HTML → plain text (inbound display) ───────────────────────────────────────

describe("htmlToText (inbound sanitiser)", () => {
	it("strips <script> blocks entirely", () => {
		expect(htmlToText("<script>alert(1)</script>hello")).toBe("hello");
	});

	it("strips <style> blocks entirely", () => {
		expect(htmlToText("<style>body{color:red}</style>hello")).toBe("hello");
	});

	it("strips HTML comments", () => {
		expect(htmlToText("<!-- secret note -->visible")).toBe("visible");
	});

	it("converts block-level tags into paragraph breaks", () => {
		expect(htmlToText("<p>one</p><p>two</p>")).toBe("one\n\ntwo");
	});

	it("decodes common HTML entities", () => {
		expect(htmlToText("&amp; &lt; &gt; &quot; &#39;")).toBe("& < > \" '");
	});

	it("decodes numeric entities (decimal + hex)", () => {
		expect(htmlToText("&#65;&#x42;")).toBe("AB");
	});

	it("collapses runs of whitespace and excess blank lines", () => {
		expect(htmlToText("a   b\n\n\n\n\nc")).toBe("a b\n\nc");
	});

	it("returns empty string for empty input", () => {
		expect(htmlToText("")).toBe("");
	});
});

// ── Outbound HTML sanitiser (defence-in-depth) ────────────────────────────────

describe("sanitizeOutboundHtml — known bypass classes are closed", () => {
	it("strips <script> blocks", () => {
		expect(sanitizeOutboundHtml("<script>x</script>safe")).toBe("safe");
	});

	it("strips <style> blocks (CSS expression / url() XSS surface)", () => {
		expect(sanitizeOutboundHtml("<style>body{x:expression(alert(1))}</style>ok")).toBe("ok");
	});

	it("strips <svg> blocks (including self-closing onload variant)", () => {
		expect(sanitizeOutboundHtml("<svg><script>x</script></svg>after")).toBe("after");
		// `<svg/onload=...>` slipped through the old `\son[a-z]+\s*=` regex
		// (no whitespace before attribute, slash separator). Now closed.
		expect(sanitizeOutboundHtml("<svg/onload=alert(1)>after")).toBe("after");
	});

	it("strips <iframe>, <object>, <embed>, <form>, <link>, <meta>, <base>", () => {
		expect(sanitizeOutboundHtml("<iframe src=evil></iframe>ok")).toBe("ok");
		expect(sanitizeOutboundHtml("<object data=evil></object>ok")).toBe("ok");
		expect(sanitizeOutboundHtml("<embed src=evil>ok")).toBe("ok");
		expect(sanitizeOutboundHtml("<form action=evil><input></form>ok")).toBe("ok");
		expect(sanitizeOutboundHtml("<link rel=stylesheet href=evil>ok")).toBe("ok");
		expect(sanitizeOutboundHtml("<meta http-equiv=refresh content=...>ok")).toBe("ok");
		expect(sanitizeOutboundHtml("<base href=https://evil>ok")).toBe("ok");
	});

	it("strips <math> / <foreignObject> XSS vectors", () => {
		expect(sanitizeOutboundHtml("<math><mtext>x</mtext></math>after")).toBe("after");
	});

	it("blocks dangerous URL schemes (javascript, vbscript, livescript, mhtml)", () => {
		const out1 = sanitizeOutboundHtml('<a href="javascript:alert(1)">x</a>');
		expect(out1).toContain("blocked:");
		expect(out1.toLowerCase()).not.toContain("javascript:");

		const out2 = sanitizeOutboundHtml('<a href="vbscript:msgbox 1">x</a>');
		expect(out2).toContain("blocked:");
		expect(out2.toLowerCase()).not.toContain("vbscript:");

		const out3 = sanitizeOutboundHtml('<a href="livescript:x">x</a>');
		expect(out3.toLowerCase()).not.toContain("livescript:");

		const out4 = sanitizeOutboundHtml('<a href="mhtml:x">x</a>');
		expect(out4.toLowerCase()).not.toContain("mhtml:");
	});

	it("normalises HTML-entity-encoded colons so the scheme check still fires", () => {
		const out = sanitizeOutboundHtml('<a href="javascript&#x3a;alert(1)">x</a>');
		// Either the scheme prefix was replaced or the entity-encoded version was stripped.
		expect(out.toLowerCase()).not.toContain("javascript:");
		expect(out).not.toContain("javascript&#x3a;");
	});

	it("strips on*= event-handler attributes (quoted, unquoted, slash-separated)", () => {
		expect(sanitizeOutboundHtml('<img src=x onerror="alert(1)">')).not.toContain("onerror");
		expect(sanitizeOutboundHtml("<img src=x onerror='alert(1)'>")).not.toContain("onerror");
		expect(sanitizeOutboundHtml("<img src=x onerror=alert(1)>")).not.toContain("onerror");
		expect(sanitizeOutboundHtml("<img/onerror=alert(1)>")).not.toContain("onerror");
	});

	it("strips style= attributes (CSS url(javascript:...) etc.)", () => {
		const out = sanitizeOutboundHtml('<p style="background:url(javascript:alert(1))">x</p>');
		expect(out).not.toMatch(/style\s*=/i);
	});

	it("strips formaction= attributes (button XSS sink)", () => {
		const out = sanitizeOutboundHtml('<button formaction="javascript:x">go</button>');
		expect(out).not.toMatch(/formaction\s*=/i);
	});

	it("closes polyglot-closure bypass via two-pass sanitisation", () => {
		// `<scr<script>ipt>` becomes `<script>` after the inner <script>...</script> is stripped.
		// Two passes catches the residue.
		const out = sanitizeOutboundHtml("<scr<script>x</script>ipt>alert(1)</script>ok");
		expect(out).not.toMatch(/<script\b/i);
	});

	it("preserves legitimate inline content", () => {
		const out = sanitizeOutboundHtml("<p>Hello <b>world</b></p>");
		expect(out).toContain("Hello");
		expect(out).toContain("<b>world</b>");
	});
});

// ── SSRF guard (URL attachment fetch) ─────────────────────────────────────────

describe("rejectIfPrivateHost — IPv4 dotted decimal", () => {
	it("blocks loopback / RFC1918 / link-local / metadata / multicast", () => {
		const ranges = [
			"127.0.0.1",
			"10.0.0.5",
			"172.16.0.1",
			"172.31.255.255",
			"192.168.1.1",
			"169.254.169.254",
			"0.0.0.0",
			"224.0.0.1",
		];
		for (const ip of ranges) {
			expect(() => rejectIfPrivateHost(ip)).toThrow(/Refusing/);
		}
	});

	it("permits public IPv4 literals", () => {
		expect(() => rejectIfPrivateHost("8.8.8.8")).not.toThrow();
		expect(() => rejectIfPrivateHost("1.1.1.1")).not.toThrow();
	});
});

describe("rejectIfPrivateHost — decimal/hex IPv4 normalisation", () => {
	it("normalises 32-bit decimal IPv4 (2130706433 = 127.0.0.1)", () => {
		expect(normaliseIpHostname("2130706433")).toBe("127.0.0.1");
		expect(() => rejectIfPrivateHost("2130706433")).toThrow(/Refusing/);
	});

	it("normalises hex IPv4 (0x7f000001 = 127.0.0.1)", () => {
		expect(normaliseIpHostname("0x7f000001")).toBe("127.0.0.1");
		expect(() => rejectIfPrivateHost("0x7f000001")).toThrow(/Refusing/);
	});

	it("normalises 32-bit decimal pointing at AWS metadata (2852039166 = 169.254.169.254)", () => {
		expect(normaliseIpHostname("2852039166")).toBe("169.254.169.254");
		expect(() => rejectIfPrivateHost("2852039166")).toThrow(/Refusing/);
	});

	it("leaves DNS names alone", () => {
		expect(normaliseIpHostname("example.com")).toBe("example.com");
	});
});

describe("rejectIfPrivateHost — IPv6 + internal hostnames", () => {
	it("blocks IPv6 loopback / ULA / link-local / IPv4-mapped", () => {
		expect(() => rejectIfPrivateHost("::1")).toThrow();
		expect(() => rejectIfPrivateHost("fc00::1")).toThrow();
		expect(() => rejectIfPrivateHost("fe80::1")).toThrow();
		expect(() => rejectIfPrivateHost("::ffff:127.0.0.1")).toThrow();
	});

	it("blocks localhost-style hostnames", () => {
		expect(() => rejectIfPrivateHost("localhost")).toThrow();
		expect(() => rejectIfPrivateHost("anything.local")).toThrow();
		expect(() => rejectIfPrivateHost("router.internal")).toThrow();
	});
});

// ── Attachment validation ─────────────────────────────────────────────────────

describe("validateAndBuildAttachments", () => {
	const TINY = btoa("hello");

	it("rejects disallowed MIME types", () => {
		expect(() =>
			validateAndBuildAttachments([
				{ name: "a.svg", content_base64: TINY, content_type: "image/svg+xml" },
			]),
		).toThrow(/MIME type not allowed/);
	});

	it("rejects executable filename extensions even with safe MIME", () => {
		expect(() =>
			validateAndBuildAttachments([
				{ name: "report.exe", content_base64: TINY, content_type: "application/pdf" },
			]),
		).toThrow(/extension not allowed/);
	});

	it("rejects inline attachment without content_id", () => {
		expect(() =>
			validateAndBuildAttachments([
				{
					name: "logo.png",
					content_base64: TINY,
					content_type: "image/png",
					is_inline: true,
				},
			]),
		).toThrow(/content_id/);
	});

	it("accepts a well-formed PNG attachment", () => {
		const built = validateAndBuildAttachments([
			{ name: "a.png", content_base64: TINY, content_type: "image/png" },
		]);
		expect(built).toHaveLength(1);
		expect(built[0]?.["@odata.type"]).toBe("#microsoft.graph.fileAttachment");
		expect(built[0]?.contentType).toBe("image/png");
	});
});

// ── Recurrence translator ─────────────────────────────────────────────────────

describe("buildRecurrence", () => {
	const start = "2026-06-01T09:00:00";
	const tz = "Pacific/Auckland";

	it("builds a daily recurrence with default interval", () => {
		const r = buildRecurrence({ pattern: "daily" }, start, tz);
		expect((r.pattern as { type: string }).type).toBe("daily");
		expect((r.pattern as { interval: number }).interval).toBe(1);
	});

	it("requires days_of_week for weekly", () => {
		expect(() => buildRecurrence({ pattern: "weekly" }, start, tz)).toThrow(/days_of_week/);
	});

	it("rejects invalid weekday names", () => {
		expect(() =>
			buildRecurrence(
				{ pattern: "weekly", days_of_week: ["munday"] },
				start,
				tz,
			),
		).toThrow(/Invalid day/);
	});

	it("requires day_of_month for monthly", () => {
		expect(() => buildRecurrence({ pattern: "monthly" }, start, tz)).toThrow(/day_of_month/);
	});

	it("requires day_of_month AND month for yearly", () => {
		expect(() =>
			buildRecurrence({ pattern: "yearly", day_of_month: 1 }, start, tz),
		).toThrow(/month/);
	});

	it("rejects fractional interval", () => {
		expect(() =>
			buildRecurrence({ pattern: "daily", interval: 1.5 }, start, tz),
		).toThrow(/positive integer/);
	});

	it("rejects both end_date and occurrences", () => {
		expect(() =>
			buildRecurrence(
				{
					pattern: "weekly",
					days_of_week: ["monday"],
					end_date: "2026-12-31",
					occurrences: 10,
				},
				start,
				tz,
			),
		).toThrow(/end_date OR.*occurrences/);
	});

	it("defaults to noEnd when neither end_date nor occurrences given", () => {
		const r = buildRecurrence(
			{ pattern: "weekly", days_of_week: ["monday"] },
			start,
			tz,
		);
		expect((r.range as { type: string }).type).toBe("noEnd");
	});
});

// ── Router smoke tests (post-toolkit-migration) ───────────────────────────────

describe("router", () => {
	const claudeOrigin = { Origin: "https://claude.ai" };

	it("answers CORS preflight from claude.ai with 204", async () => {
		const response = await SELF.fetch("https://example.com/mcp", {
			method: "OPTIONS",
			headers: {
				...claudeOrigin,
				"Access-Control-Request-Method": "POST",
				"Access-Control-Request-Headers": "Authorization,Content-Type",
			},
		});
		expect(response.status).toBe(204);
		expect(response.headers.get("Access-Control-Allow-Origin")).toBe("https://claude.ai");
	});

	it("rejects CORS preflight from disallowed origin", async () => {
		const response = await SELF.fetch("https://example.com/mcp", {
			method: "OPTIONS",
			headers: { Origin: "https://attacker.example.com" },
		});
		expect(response.status).toBe(403);
	});

	it("rejects /mcp without Bearer (401 + WWW-Authenticate)", async () => {
		const response = await SELF.fetch("https://example.com/mcp", { method: "POST" });
		expect(response.status).toBe(401);
		expect(response.headers.get("WWW-Authenticate") ?? "").toContain("Bearer");
		expect(response.headers.get("WWW-Authenticate") ?? "").toContain(
			"oauth-protected-resource",
		);
	});

	it("rejects /oauth/status without Bearer (401)", async () => {
		const response = await SELF.fetch("https://example.com/oauth/status");
		expect(response.status).toBe(401);
	});

	it("serves /oauth/start as an HTML form (public path)", async () => {
		const response = await SELF.fetch("https://example.com/oauth/start");
		expect(response.status).toBe(200);
		expect(response.headers.get("Content-Type") ?? "").toContain("text/html");
		const body = await response.text();
		expect(body).toContain("Connect Outlook");
		expect(body).toContain("MCP_APPROVAL_CODE");
	});

	it("serves OAuth discovery metadata (public)", async () => {
		const r1 = await SELF.fetch(
			"https://example.com/.well-known/oauth-authorization-server",
		);
		expect(r1.status).toBe(200);
		const meta = (await r1.json()) as {
			authorization_endpoint: string;
			token_endpoint: string;
		};
		expect(meta.authorization_endpoint).toMatch(/\/authorize$/);
		expect(meta.token_endpoint).toMatch(/\/token$/);
	});

	it("serves OAuth protected-resource metadata (public)", async () => {
		const response = await SELF.fetch(
			"https://example.com/.well-known/oauth-protected-resource",
		);
		expect(response.status).toBe(200);
		const meta = (await response.json()) as { resource: string };
		expect(meta.resource).toMatch(/\/mcp$/);
	});

	it("returns 401 (not 404) for unknown paths — bearer middleware runs first", async () => {
		const response = await SELF.fetch("https://example.com/does-not-exist");
		expect(response.status).toBe(401);
	});
});

// ── VTT transcript parser ─────────────────────────────────────────────────────
// Microsoft Graph returns text/vtt only; we parse to structured JSON locally.

describe("parseVtt (Teams transcript parser)", () => {
	it("parses a single cue with a voice tag", () => {
		const vtt = `WEBVTT

00:00:00.030 --> 00:00:02.140
<v Frank Wiess>Hey, Pavel. Hey, Mary.</v>
`;
		const { count, cues } = parseVtt(vtt);
		expect(count).toBe(1);
		expect(cues[0]).toEqual({
			start: 0.03,
			end: 2.14,
			speaker: "Frank Wiess",
			text: "Hey, Pavel. Hey, Mary.",
		});
	});

	it("parses multiple cues in order", () => {
		const vtt = `WEBVTT

00:00:00.000 --> 00:00:02.000
<v Alice>First line.</v>

00:00:02.000 --> 00:00:04.500
<v Bob>Second line.</v>
`;
		const { count, cues } = parseVtt(vtt);
		expect(count).toBe(2);
		expect(cues[0]?.speaker).toBe("Alice");
		expect(cues[1]?.speaker).toBe("Bob");
		expect(cues[1]?.end).toBe(4.5);
	});

	it("skips WEBVTT header, NOTE blocks, and cue identifiers", () => {
		const vtt = `WEBVTT - Meeting transcript

NOTE
This is a note block and should be ignored.

73f9b81b-0099-4322-bb31-ec4a3a6e23b6/1-0
00:00:01.000 --> 00:00:02.000
<v Alice>Hello.</v>
`;
		const { count, cues } = parseVtt(vtt);
		expect(count).toBe(1);
		expect(cues[0]?.text).toBe("Hello.");
		expect(cues[0]?.speaker).toBe("Alice");
	});

	it("strips inline tags (italic, color, timestamp markers) but keeps text", () => {
		const vtt = `WEBVTT

00:00:00.000 --> 00:00:05.000
<v Alice><i>emphasised</i> and <c.yellow>coloured</c> with <00:00:02.500>marker.</v>
`;
		const { cues } = parseVtt(vtt);
		expect(cues[0]?.text).toBe("emphasised and coloured with marker.");
		expect(cues[0]?.speaker).toBe("Alice");
	});

	it("handles voice tags with class names (extracts speaker, ignores classes)", () => {
		const vtt = `WEBVTT

00:00:00.000 --> 00:00:01.000
<v.first.loud Esme>Loud opener.</v>
`;
		const { cues } = parseVtt(vtt);
		expect(cues[0]?.speaker).toBe("Esme");
		expect(cues[0]?.text).toBe("Loud opener.");
	});

	it("handles negative offsets (transcription started mid-call)", () => {
		const vtt = `WEBVTT

-00:00:30.000 --> 00:00:00.000
<v Alice>Already mid-sentence...</v>
`;
		const { cues } = parseVtt(vtt);
		expect(cues[0]?.start).toBe(-30);
		expect(cues[0]?.end).toBe(0);
	});

	it("preserves text without a voice tag (speaker = null)", () => {
		const vtt = `WEBVTT

00:00:00.000 --> 00:00:01.000
No speaker attribution here.
`;
		const { cues } = parseVtt(vtt);
		expect(cues[0]?.speaker).toBeNull();
		expect(cues[0]?.text).toBe("No speaker attribution here.");
	});

	it("returns empty cues for empty or header-only input", () => {
		expect(parseVtt("").count).toBe(0);
		expect(parseVtt("WEBVTT\n").count).toBe(0);
		expect(parseVtt("WEBVTT\n\nNOTE\nJust a note.\n").count).toBe(0);
	});

	it("skips malformed cues but keeps well-formed ones around them", () => {
		const vtt = `WEBVTT

99:99:99.bad --> 00:00:01.000
This cue has a bad timing line.

00:00:02.000 --> 00:00:03.000
<v Alice>This cue is fine.</v>
`;
		const { count, cues } = parseVtt(vtt);
		expect(count).toBe(1);
		expect(cues[0]?.speaker).toBe("Alice");
	});
});

// ── Join URL extractor (body-preview fallback for events lacking structured fields) ───

describe("extractJoinUrlFromBody (Outlook add-in fallback)", () => {
	it("extracts the new short-URL Teams format from a real Outlook body preview", () => {
		const body =
			"________________________________________________________________________________\r\n" +
			"Microsoft Teams meeting\r\n" +
			"Join: https://teams.microsoft.com/meet/42948338826416?p=U3RjMQ8vO4KbJIH0po\r\n" +
			"Meeting ID: 429 483 388 264 16\r\n" +
			"Passcode: Wq2mY2VF\r\n" +
			"__________________";
		expect(extractJoinUrlFromBody(body)).toBe(
			"https://teams.microsoft.com/meet/42948338826416?p=U3RjMQ8vO4KbJIH0po",
		);
	});

	it("extracts the legacy meetup-join URL with context query string", () => {
		const body =
			"Click here to join the meeting: https://teams.microsoft.com/l/meetup-join/19%3Ameeting_abc%40thread.v2/0?context=%7B%22Tid%22%3A%22xyz%22%7D extra trailing text";
		expect(extractJoinUrlFromBody(body)).toBe(
			"https://teams.microsoft.com/l/meetup-join/19%3Ameeting_abc%40thread.v2/0?context=%7B%22Tid%22%3A%22xyz%22%7D",
		);
	});

	it("extracts the personal-account teams.live.com variant", () => {
		const body = "Join: https://teams.live.com/meet/9876543210?p=ABCdef";
		expect(extractJoinUrlFromBody(body)).toBe(
			"https://teams.live.com/meet/9876543210?p=ABCdef",
		);
	});

	it("returns null when no Teams meeting URL is present", () => {
		expect(extractJoinUrlFromBody("Some unrelated body text")).toBeNull();
		expect(
			extractJoinUrlFromBody(
				"https://teams.microsoft.com/meetingOptions/?organizerId=xyz",
			),
		).toBeNull();
		expect(extractJoinUrlFromBody("")).toBeNull();
		expect(extractJoinUrlFromBody(undefined)).toBeNull();
	});

	it("returns the first matching URL when multiple are present", () => {
		const body =
			"Quick link: https://teams.microsoft.com/meet/111?p=AAA " +
			"Backup: https://teams.microsoft.com/l/meetup-join/19%3Ameeting_xxx%40thread.v2/0?context=%7B%7D";
		expect(extractJoinUrlFromBody(body)).toBe(
			"https://teams.microsoft.com/meet/111?p=AAA",
		);
	});
});

// ── Date-only input normalisation (Microsoft OData filter inclusivity) ────────

describe("normaliseStartDate / normaliseEndDate", () => {
	it("expands a date-only start to start-of-day UTC", () => {
		expect(normaliseStartDate("2026-05-19")).toBe(
			"2026-05-19T00:00:00.000Z",
		);
	});

	it("expands a date-only end to end-of-day UTC (NOT next-day midnight)", () => {
		expect(normaliseEndDate("2026-05-19")).toBe(
			"2026-05-19T23:59:59.999Z",
		);
	});

	it("leaves full ISO datetime strings untouched", () => {
		const isoStart = "2026-05-19T02:00:00Z";
		const isoEnd = "2026-05-19T23:59:59.999Z";
		expect(normaliseStartDate(isoStart)).toBe(isoStart);
		expect(normaliseEndDate(isoEnd)).toBe(isoEnd);
	});

	it("leaves datetime strings with timezone offset untouched", () => {
		const offset = "2026-05-19T14:00:00+12:00";
		expect(normaliseStartDate(offset)).toBe(offset);
		expect(normaliseEndDate(offset)).toBe(offset);
	});

	it("doesn't false-positive on malformed strings (passes through)", () => {
		// Anything that isn't strictly YYYY-MM-DD passes through unchanged so
		// Graph itself returns a clear validation error rather than us silently
		// reshaping bad input.
		expect(normaliseStartDate("2026/05/19")).toBe("2026/05/19");
		expect(normaliseEndDate("yesterday")).toBe("yesterday");
		expect(normaliseStartDate("2026-5-19")).toBe("2026-5-19");
	});
});

// ── Reply-to-self recipient redirect ───────────────────────────────────────────

describe("computeSelfReplyRecipients", () => {
	const owner = "me@bashco.co.nz";

	const msg = (
		from: string,
		to: string[],
		cc: string[] = [],
	) => ({
		from: { emailAddress: { address: from } },
		toRecipients: to.map((a) => ({ emailAddress: { address: a } })),
		ccRecipients: cc.map((a) => ({ emailAddress: { address: a } })),
	});

	it("returns null when replying to someone else's message (Graph default is correct)", () => {
		const original = msg("adam@example.com", [owner]);
		expect(computeSelfReplyRecipients(original, owner, false)).toBeNull();
	});

	it("matches the sender case-insensitively before deciding it's a self-reply", () => {
		const original = msg("ME@BashCo.co.NZ", ["client@example.com"]);
		const result = computeSelfReplyRecipients(original, owner, false);
		expect(result).not.toBeNull();
		expect(result?.toRecipients).toEqual([
			{ emailAddress: { address: "client@example.com" } },
		]);
	});

	it("redirects a self-sent reply to the original recipients", () => {
		const original = msg(owner, ["client@example.com"]);
		const result = computeSelfReplyRecipients(original, owner, false);
		expect(result?.toRecipients).toEqual([
			{ emailAddress: { address: "client@example.com" } },
		]);
		// Plain reply drops CC.
		expect(result?.ccRecipients).toEqual([]);
	});

	it("removes the owner from the recipient list so the reply never bounces to self", () => {
		const original = msg(owner, [owner, "client@example.com"], [owner]);
		const result = computeSelfReplyRecipients(original, owner, true);
		expect(result?.toRecipients).toEqual([
			{ emailAddress: { address: "client@example.com" } },
		]);
		expect(result?.ccRecipients).toEqual([]);
	});

	it("includes CC recipients only for reply-all", () => {
		const original = msg(owner, ["client@example.com"], ["cc@example.com"]);
		expect(computeSelfReplyRecipients(original, owner, false)?.ccRecipients).toEqual([]);
		expect(computeSelfReplyRecipients(original, owner, true)?.ccRecipients).toEqual([
			{ emailAddress: { address: "cc@example.com" } },
		]);
	});

	it("dedupes repeated recipients", () => {
		const original = msg(owner, [
			"client@example.com",
			"CLIENT@example.com",
			"other@example.com",
		]);
		expect(computeSelfReplyRecipients(original, owner, false)?.toRecipients).toEqual([
			{ emailAddress: { address: "client@example.com" } },
			{ emailAddress: { address: "other@example.com" } },
		]);
	});

	it("returns an empty To: list when the owner was the only recipient (caller refuses to send)", () => {
		const original = msg(owner, [owner]);
		const result = computeSelfReplyRecipients(original, owner, false);
		expect(result).not.toBeNull();
		expect(result?.toRecipients).toEqual([]);
	});
});
