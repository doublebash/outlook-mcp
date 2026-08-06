import { describe, it, expect, vi, beforeEach } from "vitest";

// Mock the Graph client layer so we can assert exactly what the To Do tools put
// on the wire. The bug these tests guard was invisible to every unit test in the
// repo and only showed up against live Graph.
vi.mock("../src/graph.js", () => ({
	graphGet: vi.fn(),
	graphGetNextLink: vi.fn(),
	graphPost: vi.fn(),
	graphPatch: vi.fn(),
	graphDelete: vi.fn(),
	graphRequestRaw: vi.fn(),
	graphPutBinary: vi.fn(),
}));

import { graphGet, graphGetNextLink, graphPost } from "../src/graph.js";
import {
	listTaskListsImpl,
	listTasksImpl,
	createTaskImpl,
	selectTasksByStatus,
} from "../src/tools/tasks.js";

const env = {} as never;
const LIST_ID = "AAMkADNkTaskListId=";

const lists = [
	{ id: LIST_ID, displayName: "Tasks", wellknownListName: "defaultList", isShared: false },
	{ id: "second-list", displayName: "Shopping", wellknownListName: "none", isShared: false },
];

beforeEach(() => {
	vi.clearAllMocks();
	vi.mocked(graphGet).mockResolvedValue({ value: lists });
	vi.mocked(graphGetNextLink).mockResolvedValue({ value: [] });
	vi.mocked(graphPost).mockResolvedValue({ id: "new-task" });
});

// Regression guard for the actual outage. Microsoft To Do answers any query
// option on these collections with `400 invalidRequest` /
// `RequestBroker--ParseUri`, so both read tools were dead from the day they
// shipped while create_task — which never sends a query string — kept working.
describe("To Do calls carry no OData query options", () => {
	it("list_task_lists sends a bare GET", async () => {
		await listTaskListsImpl(env);

		expect(vi.mocked(graphGet).mock.calls[0]?.[1]).toBe("/me/todo/lists");
		expect(vi.mocked(graphGet).mock.calls[0]?.[2]).toBeUndefined();
	});

	it("list_tasks sends a bare GET for the tasks collection", async () => {
		vi.mocked(graphGet)
			.mockResolvedValueOnce({ value: lists })
			.mockResolvedValueOnce({ value: [] });

		await listTasksImpl(env, {});

		for (const call of vi.mocked(graphGet).mock.calls) {
			expect(call[1]).not.toContain("$");
			expect(call[2]).toBeUndefined();
		}
	});

	it("list_tasks still sends nothing when a status is asked for", async () => {
		vi.mocked(graphGet)
			.mockResolvedValueOnce({ value: lists })
			.mockResolvedValueOnce({ value: [] });

		await listTasksImpl(env, { status: "completed" });

		expect(vi.mocked(graphGet).mock.calls.at(-1)?.[2]).toBeUndefined();
	});
});

describe("status filtering happens client-side", () => {
	const tasks = [
		{ id: "1", title: "open", status: "notStarted" },
		{ id: "2", title: "doing", status: "inProgress" },
		{ id: "3", title: "done", status: "completed" },
		{ id: "4", title: "blocked", status: "waitingOnOthers" },
		{ id: "5", title: "later", status: "deferred" },
	];

	const titles = (items: unknown[]) => (items as Array<{ title?: string }>).map((t) => t.title);

	it("excludes completed tasks by default", () => {
		expect(titles(selectTasksByStatus(tasks, undefined))).toEqual([
			"open",
			"doing",
			"blocked",
			"later",
		]);
	});

	// waitingOnOthers and deferred are Graph taskStatus values the tool's own
	// enum can't ask for. They are open work, so the default view must keep them.
	it("keeps the statuses the tool enum cannot name", () => {
		expect(titles(selectTasksByStatus(tasks, undefined))).toContain("blocked");
		expect(titles(selectTasksByStatus(tasks, undefined))).toContain("later");
	});

	it("returns only the requested status when one is given", () => {
		expect(titles(selectTasksByStatus(tasks, "completed"))).toEqual(["done"]);
		expect(titles(selectTasksByStatus(tasks, "inProgress"))).toEqual(["doing"]);
	});

	it("reaches the sanitised output of list_tasks", async () => {
		vi.mocked(graphGet)
			.mockResolvedValueOnce({ value: lists })
			.mockResolvedValueOnce({ value: tasks });

		const result = (await listTasksImpl(env, {})) as Array<{ title?: string }>;

		expect(result.map((t) => t.title)).toEqual(["open", "doing", "blocked", "later"]);
	});
});

// Without a server-side $filter, completed tasks take up slots in page one, so
// stopping there can hide open tasks entirely.
describe("paging", () => {
	it("follows @odata.nextLink instead of truncating at the first page", async () => {
		vi.mocked(graphGet)
			.mockResolvedValueOnce({ value: lists })
			.mockResolvedValueOnce({
				value: [{ id: "1", title: "page one", status: "notStarted" }],
				"@odata.nextLink": "https://graph.microsoft.com/v1.0/me/todo/lists/X/tasks?$skiptoken=A",
			});
		vi.mocked(graphGetNextLink).mockResolvedValueOnce({
			value: [{ id: "2", title: "page two", status: "notStarted" }],
		});

		const result = (await listTasksImpl(env, {})) as Array<{ title?: string }>;

		expect(result.map((t) => t.title)).toEqual(["page one", "page two"]);
	});

	it("stops at a bounded number of pages rather than looping forever", async () => {
		vi.mocked(graphGet).mockResolvedValueOnce({ value: lists }).mockResolvedValueOnce({
			value: [],
			"@odata.nextLink": "https://graph.microsoft.com/v1.0/me/todo/lists/X/tasks?$skiptoken=A",
		});
		vi.mocked(graphGetNextLink).mockResolvedValue({
			value: [],
			"@odata.nextLink": "https://graph.microsoft.com/v1.0/me/todo/lists/X/tasks?$skiptoken=B",
		});

		await listTasksImpl(env, {});

		expect(vi.mocked(graphGetNextLink).mock.calls.length).toBeLessThan(20);
	});
});

describe("default list resolution", () => {
	it("picks the wellknown defaultList rather than the first list returned", async () => {
		vi.mocked(graphGet)
			.mockResolvedValueOnce({
				value: [
					{ id: "shopping", wellknownListName: "none" },
					{ id: LIST_ID, wellknownListName: "defaultList" },
				],
			})
			.mockResolvedValueOnce({ value: [] });

		await listTasksImpl(env, {});

		expect(vi.mocked(graphGet).mock.calls.at(-1)?.[1]).toBe(`/me/todo/lists/${LIST_ID}/tasks`);
	});

	it("skips the lookup entirely when an explicit list_id is given", async () => {
		vi.mocked(graphGet).mockResolvedValueOnce({ value: [] });

		await listTasksImpl(env, { list_id: "explicit-list" });

		expect(vi.mocked(graphGet).mock.calls).toHaveLength(1);
		expect(vi.mocked(graphGet).mock.calls[0]?.[1]).toBe("/me/todo/lists/explicit-list/tasks");
	});
});

// create_task was the one To Do tool that never broke. Keep it that way.
describe("create_task", () => {
	it("posts to the resolved default list with no query options", async () => {
		await createTaskImpl(env, { title: "Write it down" });

		expect(vi.mocked(graphPost).mock.calls[0]?.[1]).toBe(`/me/todo/lists/${LIST_ID}/tasks`);
		expect(vi.mocked(graphPost).mock.calls[0]?.[2]).toMatchObject({
			title: "Write it down",
			importance: "normal",
		});
	});
});
