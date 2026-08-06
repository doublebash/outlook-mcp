import { describe, it, expect } from "vitest";

import { buildGraphPath } from "../src/graph.js";

// Microsoft Graph is served by several different backends behind one hostname,
// and they don't agree on how tolerant they are about the query string.
//
// `URLSearchParams` percent-encodes `$` to `%24` (it only leaves ASCII
// alphanumerics and `*-._` alone). The Outlook/Exchange backends that serve
// /me/messages, /me/events and /me/contacts quietly decode that and carry on,
// which is why every other tool in this worker worked while sending `%24select`.
//
// The Microsoft To Do backend does not. `/me/todo/*` hands the URI to a request
// broker that never decodes the parameter *name*, fails to parse the URI, and
// returns a flat `400 invalidRequest` with `innerError.code` of
// `RequestBroker--ParseUri`. That took out every To Do call carrying a query
// string — list_task_lists and list_tasks — while create_task kept working,
// because its only GET (resolving the default list) sends no query at all.
//
// So query strings are built here instead: OData option names stay literal, the
// way every example in Microsoft's own docs writes them.
describe("buildGraphPath — OData option names", () => {
	it("leaves the $ of an option name literal rather than encoding it to %24", () => {
		expect(buildGraphPath("/me/todo/lists", { $select: "id,displayName" })).toBe(
			"/me/todo/lists?$select=id,displayName",
		);
	});

	it("builds the To Do task query that was returning RequestBroker--ParseUri", () => {
		expect(
			buildGraphPath("/me/todo/lists/ABC/tasks", {
				$select: "id,title,status",
				$filter: "status ne 'completed'",
			}),
		).toBe("/me/todo/lists/ABC/tasks?$select=id,title,status&$filter=status%20ne%20'completed'");
	});
});

describe("buildGraphPath — value encoding", () => {
	it("keeps the commas separating $select fields readable, not %2C", () => {
		expect(buildGraphPath("/me/messages", { $select: "id,subject,from" })).toBe(
			"/me/messages?$select=id,subject,from",
		);
	});

	it("keeps the / of a nested $orderby property and encodes the space as %20, not +", () => {
		expect(buildGraphPath("/me/events", { $orderby: "start/dateTime desc" })).toBe(
			"/me/events?$orderby=start/dateTime%20desc",
		);
	});

	// A `&` or `?` arriving inside a user-supplied search term must not be able to
	// open a new query parameter.
	it("percent-encodes characters that would otherwise split the query string", () => {
		expect(buildGraphPath("/me/contacts", { $filter: "contains(displayName,'A&B?c=d')" })).toBe(
			"/me/contacts?$filter=contains(displayName,'A%26B%3Fc%3Dd')",
		);
	});

	it("serialises numeric values", () => {
		expect(buildGraphPath("/me/messages", { $top: 50 })).toBe("/me/messages?$top=50");
	});
});

describe("buildGraphPath — omitted values", () => {
	it("returns the path untouched when there is no query", () => {
		expect(buildGraphPath("/me/todo/lists")).toBe("/me/todo/lists");
	});

	it("returns the path untouched when every value is absent", () => {
		expect(buildGraphPath("/me/todo/lists", { $select: undefined, $filter: "" })).toBe(
			"/me/todo/lists",
		);
	});

	it("drops only the absent parameters", () => {
		expect(buildGraphPath("/me/messages", { $select: "id", $filter: undefined, $top: 10 })).toBe(
			"/me/messages?$select=id&$top=10",
		);
	});
});
