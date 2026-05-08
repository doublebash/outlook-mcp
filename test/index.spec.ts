import { describe, it, expect } from "vitest";
import { SELF } from "cloudflare:test";
import { htmlToText } from "../src/sanitize";

// ── Pure-function tests: HTML sanitization ────────────────────────────────────

describe("htmlToText", () => {
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
		// Both opening and closing <p> emit a newline, producing a paragraph break.
		expect(htmlToText("<p>one</p><p>two</p>")).toBe("one\n\ntwo");
	});

	it("decodes common HTML entities", () => {
		expect(htmlToText("&amp; &lt; &gt; &quot; &#39;")).toBe("& < > \" '");
	});

	it("collapses runs of whitespace and excess blank lines", () => {
		expect(htmlToText("a   b\n\n\n\n\nc")).toBe("a b\n\nc");
	});

	it("returns empty string for empty input", () => {
		expect(htmlToText("")).toBe("");
	});
});

// ── Router behavior: paths that don't require Bearer auth or KV state ─────────

describe("router", () => {
	it("answers OPTIONS preflight with 204 and CORS headers", async () => {
		const response = await SELF.fetch("https://example.com/mcp", {
			method: "OPTIONS",
		});
		expect(response.status).toBe(204);
		expect(response.headers.get("Access-Control-Allow-Origin")).toBe("*");
		expect(response.headers.get("Access-Control-Allow-Methods")).toContain("POST");
	});

	it("returns 404 with CORS headers for unknown paths", async () => {
		const response = await SELF.fetch("https://example.com/does-not-exist");
		expect(response.status).toBe(404);
		expect(response.headers.get("Access-Control-Allow-Origin")).toBe("*");
	});

	it("returns 404 for GET /mcp (only POST is supported)", async () => {
		const response = await SELF.fetch("https://example.com/mcp");
		expect(response.status).toBe(404);
	});

	it("rejects /oauth/status without Bearer token (401)", async () => {
		const response = await SELF.fetch("https://example.com/oauth/status");
		expect(response.status).toBe(401);
	});

	it("serves /oauth/start as an HTML form", async () => {
		const response = await SELF.fetch("https://example.com/oauth/start");
		expect(response.status).toBe(200);
		expect(response.headers.get("Content-Type") ?? "").toContain("text/html");
		const body = await response.text();
		expect(body).toContain("Connect Outlook");
	});
});
