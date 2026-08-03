import { describe, it, expect } from "vitest";
import { ToolError } from "@bashco/mcp-toolkit";

import { resolveNextLinkPath } from "../src/graph.js";

// `@odata.nextLink` arrives inside a Graph response body and is then fetched
// with the mailbox bearer token attached, so it gets the same treatment as any
// other URL that came from upstream: follow it whole, but only after confirming
// it still points at Graph.
describe("resolveNextLinkPath — accepted links", () => {
	it("returns the path and query below the API base", () => {
		expect(
			resolveNextLinkPath("https://graph.microsoft.com/v1.0/me/events?$top=200&$skip=200"),
		).toBe("/me/events?$top=200&$skip=200");
	});

	it("preserves a $skiptoken untouched, opaque padding and all", () => {
		expect(
			resolveNextLinkPath(
				"https://graph.microsoft.com/v1.0/me/messages?$skiptoken=RFNwdAIAAQ%3D%3D",
			),
		).toBe("/me/messages?$skiptoken=RFNwdAIAAQ%3D%3D");
	});

	it("keeps every other query parameter Graph encoded into the link", () => {
		const resolved = resolveNextLinkPath(
			"https://graph.microsoft.com/v1.0/me/events" +
				"?$select=id%2Csubject&$orderby=start%2FdateTime%20desc&$skiptoken=X",
		);
		expect(resolved).toContain("$select=id%2Csubject");
		expect(resolved).toContain("$orderby=start%2FdateTime%20desc");
		expect(resolved).toContain("$skiptoken=X");
	});

	it("handles a link with no query at all", () => {
		expect(resolveNextLinkPath("https://graph.microsoft.com/v1.0/me/contacts")).toBe(
			"/me/contacts",
		);
	});
});

describe("resolveNextLinkPath — rejected links", () => {
	const rejects = (link: string) => {
		expect(() => resolveNextLinkPath(link)).toThrow(ToolError);
	};

	it("rejects a different host", () => {
		rejects("https://evil.example.com/v1.0/me/events?$skip=200");
	});

	it("rejects a lookalike host", () => {
		rejects("https://graph.microsoft.com.evil.example/v1.0/me/events");
	});

	it("rejects plain http on the right host", () => {
		rejects("http://graph.microsoft.com/v1.0/me/events");
	});

	it("rejects another API version outside the configured base", () => {
		rejects("https://graph.microsoft.com/beta/me/events?$skip=200");
	});

	// "/v1.0evil/..." shares a string prefix with "/v1.0" but is a different
	// path segment — a prefix test alone would wave it through.
	it("rejects a path that only prefix-matches the base", () => {
		rejects("https://graph.microsoft.com/v1.0evil/me/events");
	});

	it("rejects a malformed URL", () => {
		rejects("not a url");
	});

	it("rejects a relative link", () => {
		rejects("/v1.0/me/events?$skip=200");
	});
});
