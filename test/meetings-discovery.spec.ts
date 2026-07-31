import { describe, it, expect, vi, beforeEach } from "vitest";

// Mock the Graph client layer so we can assert exactly what the discovery
// query asks for. The `isOnlineMeeting eq true` filter below was rejected by
// Graph on every call, which only showed up against a live tenant.
vi.mock("../src/graph.js", () => ({
	graphGet: vi.fn(),
	graphPost: vi.fn(),
	graphPatch: vi.fn(),
	graphDelete: vi.fn(),
	graphRequestRaw: vi.fn(),
	graphPutBinary: vi.fn(),
}));

import { graphGet } from "../src/graph.js";
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

// Routes the three call shapes the impl makes: the events page, the
// onlineMeetings JoinWebUrl lookup, then recordings/transcripts per meeting.
function routeGraph(
	eventPages: Array<{ value: unknown[]; "@odata.nextLink"?: string }>,
	opts: { withContent?: boolean } = {},
) {
	let page = 0;
	vi.mocked(graphGet).mockImplementation(
		async (_e: unknown, path: string): Promise<unknown> => {
			if (path === "/me/events") return eventPages[page++] ?? { value: [] };
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
}

function eventsQueries(): Array<Record<string, string | number>> {
	return vi
		.mocked(graphGet)
		.mock.calls.filter(c => c[1] === "/me/events")
		.map(c => c[2] as Record<string, string | number>);
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

		const queries = eventsQueries();
		expect(queries).toHaveLength(2);
		expect(queries[1]?.$skiptoken).toBe("PAGE2");
		expect(result.events_scanned).toBe(3);
		expect(result.count).toBe(1);
	});

	it("carries the filter and select onto later pages", async () => {
		routeGraph([
			{
				value: [plainEvent("p1")],
				"@odata.nextLink":
					"https://graph.microsoft.com/v1.0/me/events?$skiptoken=PAGE2",
			},
			{ value: [] },
		]);

		await listRecentMeetingRecordingsImpl(env, {});

		const [first, second] = eventsQueries();
		expect(second?.$filter).toBe(first?.$filter);
		expect(second?.$select).toBe(first?.$select);
	});

	it("stops when Graph returns no nextLink", async () => {
		routeGraph([{ value: [plainEvent("p1")] }]);

		await listRecentMeetingRecordingsImpl(env, {});

		expect(eventsQueries()).toHaveLength(1);
	});

	it("stops rather than looping when a nextLink carries no skiptoken", async () => {
		routeGraph([
			{
				value: [plainEvent("p1")],
				"@odata.nextLink": "https://graph.microsoft.com/v1.0/me/events",
			},
			{ value: [teamsEvent("t1")] },
		]);

		await listRecentMeetingRecordingsImpl(env, {});

		expect(eventsQueries()).toHaveLength(1);
	});

	it("caps paging so a long window cannot run unbounded", async () => {
		// Every page is non-Teams and always advertises another page.
		vi.mocked(graphGet).mockImplementation(
			async (_e: unknown, path: string): Promise<unknown> => {
				if (path === "/me/events") {
					return {
						value: [plainEvent("p")],
						"@odata.nextLink":
							"https://graph.microsoft.com/v1.0/me/events?$skiptoken=MORE",
					};
				}
				return { value: [] };
			},
		);

		await listRecentMeetingRecordingsImpl(env, { within_days: 90 });

		expect(eventsQueries().length).toBeLessThanOrEqual(10);
	});
});
