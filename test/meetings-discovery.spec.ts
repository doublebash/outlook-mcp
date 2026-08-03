import { describe, it, expect, vi, beforeEach } from "vitest";

// Mock the Graph client layer so we can assert exactly what the discovery
// query asks for. The `isOnlineMeeting eq true` filter below was rejected by
// Graph on every call, which only showed up against a live tenant.
vi.mock("../src/graph.js", () => ({
	graphGet: vi.fn(),
	graphGetNextLink: vi.fn(),
	graphPost: vi.fn(),
	graphPatch: vi.fn(),
	graphDelete: vi.fn(),
	graphRequestRaw: vi.fn(),
	graphPutBinary: vi.fn(),
}));

import { graphGet, graphGetNextLink } from "../src/graph.js";
import { listRecentMeetingRecordingsImpl } from "../src/tools/meetings.js";

const env = {} as never;

function teamsEvent(id: string) {
	return {
		id,
		subject: `Meeting ${id}`,
		start: { dateTime: "2026-07-28T01:00:00.0000000" },
		end: { dateTime: "2026-07-28T02:00:00.0000000" },
		isOnlineMeeting: true,
		onlineMeetingUrl: `https://teams.microsoft.com/meet/${id}`,
		bodyPreview: "",
		organizer: { emailAddress: { name: "Jerome Jacobs" } },
	};
}

function plainEvent(id: string) {
	return {
		id,
		subject: `Blocked out ${id}`,
		start: { dateTime: "2026-07-28T01:00:00.0000000" },
		end: { dateTime: "2026-07-28T02:00:00.0000000" },
		isOnlineMeeting: false,
	};
}

// Routes the call shapes the impl makes: the first events page (graphGet), each
// subsequent page (graphGetNextLink), the onlineMeetings JoinWebUrl lookup, then
// recordings/transcripts per meeting.
//
// Pages after the first deliberately arrive through a different function now —
// the impl follows Graph's nextLink URL rather than rebuilding the query, so the
// two paths are worth keeping visibly distinct in the harness.
function routeGraph(
	eventPages: Array<{ value: unknown[]; "@odata.nextLink"?: string }>,
	opts: { withContent?: boolean } = {},
) {
	let page = 0;
	const nextPage = () => eventPages[page++] ?? { value: [] };

	vi.mocked(graphGet).mockImplementation(
		async (_e: unknown, path: string): Promise<unknown> => {
			if (path === "/me/events") return nextPage();
			if (path === "/me/onlineMeetings") return { value: [{ id: "om-1" }] };
			if (path.endsWith("/recordings")) {
				return {
					value: opts.withContent
						? [{ id: "rec-1", createdDateTime: "2026-07-28T02:05:00Z" }]
						: [],
				};
			}
			if (path.endsWith("/transcripts")) {
				return {
					value: opts.withContent
						? [{ id: "tr-1", createdDateTime: "2026-07-28T02:06:00Z" }]
						: [],
				};
			}
			return { value: [] };
		},
	);

	vi.mocked(graphGetNextLink).mockImplementation(async (): Promise<unknown> => nextPage());
}

function eventsQueries(): Array<Record<string, string | number>> {
	return vi
		.mocked(graphGet)
		.mock.calls.filter(c => c[1] === "/me/events")
		.map(c => c[2] as Record<string, string | number>);
}

// The exact nextLink URLs handed to the follower, in order.
function nextLinksFollowed(): string[] {
	return vi.mocked(graphGetNextLink).mock.calls.map(c => c[1] as string);
}

// Total requests for a page of events, however they were routed.
function eventPageRequests(): number {
	return eventsQueries().length + nextLinksFollowed().length;
}

beforeEach(() => {
	vi.clearAllMocks();
});

describe("list_recent_meeting_recordings — event query", () => {
	// Regression guard: Graph rejects the whole request with
	// `ErrorInvalidProperty: The property 'isOnlineMeeting' does not support
	// filtering`, so this must never go back into $filter.
	it("does not filter on isOnlineMeeting", async () => {
		routeGraph([{ value: [] }]);

		await listRecentMeetingRecordingsImpl(env, {});

		const [query] = eventsQueries();
		expect(query.$filter).not.toContain("isOnlineMeeting");
	});

	it("still bounds the query by the date window", async () => {
		routeGraph([{ value: [] }]);

		await listRecentMeetingRecordingsImpl(env, { within_days: 30 });

		const [query] = eventsQueries();
		expect(query.$filter).toMatch(/start\/dateTime ge '.+' and start\/dateTime le '.+'/);
		expect(query.$select).toContain("isOnlineMeeting");
	});

	it("keeps non-Teams events out of the results", async () => {
		routeGraph([{ value: [plainEvent("p1"), teamsEvent("t1"), plainEvent("p2")] }], {
			withContent: true,
		});

		const result = (await listRecentMeetingRecordingsImpl(env, {})) as {
			events_scanned: number;
			online_meetings: number;
			count: number;
		};

		expect(result.events_scanned).toBe(3);
		expect(result.online_meetings).toBe(1);
		expect(result.count).toBe(1);
	});

	it("omits meetings that have neither a recording nor a transcript", async () => {
		routeGraph([{ value: [teamsEvent("t1")] }], { withContent: false });

		const result = (await listRecentMeetingRecordingsImpl(env, {})) as {
			online_meetings: number;
			count: number;
		};

		expect(result.online_meetings).toBe(1);
		expect(result.count).toBe(0);
	});
});

describe("list_recent_meeting_recordings — paging", () => {
	// With the online-meeting test moved client-side, $top bounds raw events
	// rather than Teams ones — so a page of non-Teams events must not hide the
	// real meetings sitting behind it.
	it("follows nextLink when the first page holds no Teams events", async () => {
		routeGraph(
			[
				{
					value: [plainEvent("p1"), plainEvent("p2")],
					"@odata.nextLink":
						"https://graph.microsoft.com/v1.0/me/events?$skiptoken=PAGE2",
				},
				{ value: [teamsEvent("t1")] },
			],
			{ withContent: true },
		);

		const result = (await listRecentMeetingRecordingsImpl(env, {})) as {
			events_scanned: number;
			count: number;
		};

		expect(eventsQueries()).toHaveLength(1);
		expect(nextLinksFollowed()).toEqual([
			"https://graph.microsoft.com/v1.0/me/events?$skiptoken=PAGE2",
		]);
		expect(result.events_scanned).toBe(3);
		expect(result.count).toBe(1);
	});

	// Regression guard. Graph pages Outlook collections with `$skip` on some
	// endpoints and `$skiptoken` on others, and its docs say not to pick either
	// value out and reuse it. An earlier version read `$skiptoken` and bailed when
	// it was absent, which turned a `$skip` nextLink into a silent one-page result
	// — no error, just missing meetings.
	it("pages on a nextLink that uses $skip rather than $skiptoken", async () => {
		routeGraph(
			[
				{
					value: [plainEvent("p1")],
					"@odata.nextLink":
						"https://graph.microsoft.com/v1.0/me/events?$top=200&$skip=200",
				},
				{ value: [teamsEvent("t1")] },
			],
			{ withContent: true },
		);

		const result = (await listRecentMeetingRecordingsImpl(env, {})) as {
			events_scanned: number;
			online_meetings: number;
			count: number;
		};

		expect(nextLinksFollowed()).toEqual([
			"https://graph.microsoft.com/v1.0/me/events?$top=200&$skip=200",
		]);
		expect(result.events_scanned).toBe(2);
		expect(result.online_meetings).toBe(1);
		expect(result.count).toBe(1);
	});

	// Graph re-encodes the original request's parameters into the nextLink, so
	// the link is the whole request — it must go back untouched rather than being
	// taken apart and reassembled.
	it("follows the nextLink verbatim instead of rebuilding the query", async () => {
		const link =
			"https://graph.microsoft.com/v1.0/me/events" +
			"?$select=id%2Csubject&$top=200&$orderby=start%2FdateTime%20desc&$skiptoken=OPAQUE%3D%3D";
		routeGraph([
			{ value: [plainEvent("p1")], "@odata.nextLink": link },
			{ value: [] },
		]);

		await listRecentMeetingRecordingsImpl(env, {});

		expect(nextLinksFollowed()).toEqual([link]);
		// The follow-up must not go back through the query-building path.
		expect(eventsQueries()).toHaveLength(1);
	});

	it("stops when Graph returns no nextLink", async () => {
		routeGraph([{ value: [plainEvent("p1")] }]);

		await listRecentMeetingRecordingsImpl(env, {});

		expect(eventPageRequests()).toBe(1);
		expect(nextLinksFollowed()).toEqual([]);
	});

	it("caps paging so a long window cannot run unbounded", async () => {
		// Every page is non-Teams and always advertises another page.
		const endlessPage = {
			value: [plainEvent("p")],
			"@odata.nextLink": "https://graph.microsoft.com/v1.0/me/events?$skiptoken=MORE",
		};
		vi.mocked(graphGet).mockImplementation(
			async (_e: unknown, path: string): Promise<unknown> =>
				path === "/me/events" ? endlessPage : { value: [] },
		);
		vi.mocked(graphGetNextLink).mockImplementation(async (): Promise<unknown> => endlessPage);

		await listRecentMeetingRecordingsImpl(env, { within_days: 90 });

		expect(eventPageRequests()).toBeLessThanOrEqual(10);
	});
});
