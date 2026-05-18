import { defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import { graphGet, graphPost } from "../graph.js";
import { sanitizeTaskList, sanitizeTaskLists } from "../sanitize.js";
import type { Env } from "../types.js";

const taskListIdSchema = z.string().min(1).max(512);

async function listTaskListsImpl(env: Env): Promise<unknown> {
	const data = (await graphGet(env, "/me/todo/lists", {
		$select: "id,displayName,isOwner,isShared,wellknownListName",
	})) as { value: unknown[] };
	return sanitizeTaskLists(data.value);
}

async function resolveTaskListId(env: Env, listId: string | undefined): Promise<string> {
	if (listId) return listId;
	const data = (await graphGet(env, "/me/todo/lists")) as {
		value: Array<{ id: string; wellknownListName?: string }>;
	};
	const def = data.value.find((l) => l.wellknownListName === "defaultList");
	const fallback = data.value[0];
	if (def) return def.id;
	if (fallback) return fallback.id;
	throw new Error("No To Do lists found on this account.");
}

async function listTasksImpl(
	env: Env,
	args: {
		list_id?: string;
		status?: "notStarted" | "inProgress" | "completed";
	},
): Promise<unknown> {
	const listId = await resolveTaskListId(env, args.list_id);
	const query: Record<string, string | number> = {
		$select: "id,title,status,importance,dueDateTime,body,createdDateTime",
	};
	if (args.status) {
		query.$filter = `status eq '${args.status}'`;
	} else {
		query.$filter = "status ne 'completed'";
	}
	const data = (await graphGet(
		env,
		`/me/todo/lists/${listId}/tasks`,
		query,
	)) as { value: unknown[] };
	return sanitizeTaskList(data.value);
}

async function createTaskImpl(
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
