import { ToolError } from "@bashco/mcp-toolkit";
import type { Env } from "../types.js";
import { calendarTools } from "./calendar.js";
import { contactsTools } from "./contacts.js";
import { emailTools } from "./email.js";
import { filesTools } from "./files.js";
import { mailboxTools } from "./mailbox.js";
import { meetingsTools } from "./meetings.js";
import { tasksTools } from "./tasks.js";

const ALL_TOOL_BUNDLES = [
	emailTools,
	calendarTools,
	contactsTools,
	tasksTools,
	mailboxTools,
	filesTools,
	meetingsTools,
] as const;

// Concatenate every per-domain bundle's tool definitions into one array.
export const toolDefinitions = ALL_TOOL_BUNDLES.flatMap((b) => b.toolDefinitions);

// Build a name -> dispatch function map for O(1) lookup.
const dispatchByName = new Map<string, (env: Env, args: unknown) => Promise<unknown>>();
for (const bundle of ALL_TOOL_BUNDLES) {
	for (const def of bundle.toolDefinitions) {
		dispatchByName.set(def.name, (env, args) => bundle.dispatch(env, def.name, args));
	}
}

export async function dispatchToolCall(
	env: Env,
	name: string,
	args: unknown,
): Promise<unknown> {
	const dispatch = dispatchByName.get(name);
	if (!dispatch) throw ToolError.validation(`unknown tool: ${name}`);
	return dispatch(env, args);
}
