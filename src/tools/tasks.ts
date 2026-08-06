import { defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import { graphGet, graphGetNextLink, graphPost } from "../graph.js";
import { sanitizeTaskList, sanitizeTaskLists } from "../sanitize.js";
import type { Env } from "../types.js";

const taskListIdSchema = z.string().min(1).max(512);

// ── Why these calls carry no query string ─────────────────────────────────────
// Microsoft To Do is the one Graph workload here that refuses OData query
// options on its collection endpoints. `/me/todo/lists?$select=...` answers
// `400 invalidRequest` with an `innerError.code` of `RequestBroker--ParseUri`
// and a message of just "Invalid request" — it names neither the option nor a
// property, so it reads like a malformed URL rather than an unsupported feature.
//
// It is neither. Every property these tools used to select is real and spelled
// as the todoTaskList and todoTask resources document it, and the rejection
// survives sending `$select` literally instead of percent-encoded as `%24select`
// (the encoding Graph's Outlook backends accept for mail, calendar, contacts and
// files). The To Do backend just doesn't implement the option; Microsoft's own
// reference hedges with "supports some of the OData query parameters" and
// declines to say which.
//
// That took out both read paths from the day they were written. create_task
// escaped it by accident: resolving the default list is a bare GET and the
// create itself is a POST, so neither carries a query string — which is exactly
// the shape proven to work, and the shape everything below now uses.
//
// The sanitisers already project down to the fields we return, so dropping
// $select costs response size and nothing else. Status filtering moves into JS.

// Graph pages To Do collections with `@odata.nextLink`. Without a server-side
// `$filter` the completed tasks now take up page slots, so an open task can sit
// behind a page boundary — follow the link rather than truncating at page one.
const MAX_PAGES = 10;

interface GraphPage {
	value?: unknown[];
	"@odata.nextLink"?: string;
}

async function collectPages(env: Env, path: string): Promise<unknown[]> {
	let page = (await graphGet(env, path)) as GraphPage;
	const items: unknown[] = [...(page.value ?? [])];

	for (let fetched = 1; fetched < MAX_PAGES; fetched += 1) {
		const next = page["@odata.nextLink"];
		if (!next) break;
		page = (await graphGetNextLink(env, next)) as GraphPage;
		items.push(...(page.value ?? []));
	}

	return items;
}

export async function listTaskListsImpl(env: Env): Promise<unknown> {
	return sanitizeTaskLists(await collectPages(env, "/me/todo/lists"));
}

async function resolveTaskListId(env: Env, listId: string | undefined): Promise<string> {
	if (listId) return listId;
	const lists = (await collectPages(env, "/me/todo/lists")) as Array<{
		id: string;
		wellknownListName?: string;
	}>;
	const def = lists.find((l) => l.wellknownListName === "defaultList");
	const fallback = lists[0];
	if (def) return def.id;
	if (fallback) return fallback.id;
	throw new Error("No To Do lists found on this account.");
}

// Graph's taskStatus enum also carries `waitingOnOthers` and `deferred`, which
// the tool's own enum doesn't offer. Both are open work, so the default view
// keeps them by excluding only `completed` rather than listing what to include.
export function selectTasksByStatus(tasks: unknown[], status: string | undefined): unknown[] {
	return (tasks as Array<{ status?: string }>).filter((task) =>
		status ? task.status === status : task.status !== "completed",
	);
}

export async function listTasksImpl(
	env: Env,
	args: {
		list_id?: string;
		status?: "notStarted" | "inProgress" | "completed";
	},
): Promise<unknown> {
	const listId = await resolveTaskListId(env, args.list_id);
	const tasks = await collectPages(env, `/me/todo/lists/${listId}/tasks`);
	return sanitizeTaskList(selectTasksByStatus(tasks, args.status));
}

export async function createTaskImpl(
	env: Env,
	args: {
		title: string;
		list_id?: string;
		due_date?: string;
		body?: string;
		importance?: "low" | "normal" | "high";
	},
): Promise<unknown> {
	const listId = await resolveTaskListId(env, args.list_id);

	const task: Record<string, unknown> = {
		title: args.title,
		importance: args.importance ?? "normal",
	};

	if (args.due_date) task.dueDateTime = { dateTime: args.due_date, timeZone: "UTC" };
	if (args.body) task.body = { content: args.body, contentType: "text" };

	const data = await graphPost(env, `/me/todo/lists/${listId}/tasks`, task);
	return { success: true, message: "Task created.", task: data };
}

export const tasksTools = defineTools<Env>({
	list_task_lists: {
		description:
			"List all Microsoft To Do task lists. Use this to find list IDs before listing or creating tasks.",
		schema: z.object({}),
		handler: (env) => listTaskListsImpl(env),
	},

	list_tasks: {
		description:
			"List tasks in a To Do task list. Defaults to the default task list and excludes completed tasks.",
		schema: z.object({
			list_id: taskListIdSchema.optional(),
			status: z.enum(["notStarted", "inProgress", "completed"]).optional(),
		}),
		handler: (env, args) => listTasksImpl(env, args),
	},

	create_task: {
		description: "Create a new task in a Microsoft To Do list.",
		schema: z.object({
			title: z.string().min(1).max(256),
			list_id: taskListIdSchema.optional(),
			due_date: z.string().min(1).max(64).optional(),
			body: z.string().max(64 * 1024).optional(),
			importance: z.enum(["low", "normal", "high"]).optional(),
		}),
		handler: (env, args) => createTaskImpl(env, args),
	},
});
