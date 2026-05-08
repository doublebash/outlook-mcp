import { graphRequest } from './graph';
import { getValidAccessToken } from './auth';
import {
  sanitizeEmailList,
  sanitizeEmailFull,
  sanitizeEventList,
  sanitizeContactList,
  sanitizeTaskList,
  sanitizeTaskLists,
  sanitizeMailboxSettings,
  sanitizeFileList,
} from './sanitize';
import {
  buildMessageBody,
  buildRecipients,
  resolveAndValidateAttachments,
} from './email-helpers';
import { buildRecurrence, type RecurrenceInput } from './calendar-helpers';
import type { Env } from './types';

// Shared description for body_type — kept identical across all email tools
// so Claude treats them the same way.
const BODY_TYPE_DESC = 'Body format: "text" (default) or "html". When "html", basic rich-text tags are allowed (b, i, u, p, br, ul, ol, li, a, img, table, etc.). Script and event-handler tags are stripped before sending.';
const ATTACHMENTS_DESC = 'Attachments array. Each item must have a name plus EXACTLY ONE source: (a) content_base64 + content_type for inline base64 bytes, OR (b) onedrive_path (e.g. "Documents/report.pdf") to fetch from the user\'s OneDrive server-side, OR (c) url (https only) to fetch from a public web URL server-side. Optional: is_inline + content_id for inline images (reference them in HTML body as cid:THE_ID). Max 3 MB per file, 10 MB total. Allowed types: pdf, docx, xlsx, pptx, png, jpg, gif, webp, txt, csv, html, json, zip. PREFER onedrive_path or url over content_base64 — the bytes never travel through the conversation, which is faster and cheaper.';

// ── Types ──────────────────────────────────────────────────────────────────────

export type ToolArgs = Record<string, unknown>;

export interface ToolDefinition {
  name: string;
  description: string;
  inputSchema: {
    type: 'object';
    properties: Record<string, unknown>;
    required?: string[];
  };
}

// ── Tool schemas ───────────────────────────────────────────────────────────────
// This is what Claude sees when it asks "what tools are available?"

export const TOOLS: ToolDefinition[] = [
  // ── Email ──────────────────────────────────────────────────────────────────
  {
    name: 'list_emails',
    description: 'Retrieve recent emails. Supports folder selection and keyword search.',
    inputSchema: {
      type: 'object',
      properties: {
        count: { type: 'number', description: 'Number of emails to retrieve (default: 20, max: 50)' },
        folder: { type: 'string', description: 'Folder to read from: inbox, sentitems, drafts, deleteditems, or a folder ID. Default: inbox' },
        search: { type: 'string', description: 'Keyword search query' },
      },
      required: [],
    },
  },
  {
    name: 'read_email',
    description: 'Read the full content of an email by its ID.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'The email ID' },
      },
      required: ['id'],
    },
  },
  {
    name: 'send_email',
    description: 'Send a new email. Supports plain text or HTML body, attachments, and inline images.',
    inputSchema: {
      type: 'object',
      properties: {
        to: { type: 'string', description: 'Recipient email address. Separate multiple with commas.' },
        subject: { type: 'string', description: 'Email subject' },
        body: { type: 'string', description: 'Email body. Plain text by default; set body_type to "html" for rich text.' },
        body_type: { type: 'string', description: BODY_TYPE_DESC },
        cc: { type: 'string', description: 'CC recipients, comma-separated (optional)' },
        bcc: { type: 'string', description: 'BCC recipients, comma-separated (optional)' },
        attachments: { type: 'array', description: ATTACHMENTS_DESC },
      },
      required: ['to', 'subject', 'body'],
    },
  },
  {
    name: 'reply_to_email',
    description: 'Reply to an email. Supports plain text or HTML, attachments, and inline images.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'The email ID to reply to' },
        body: { type: 'string', description: 'Reply message body' },
        body_type: { type: 'string', description: BODY_TYPE_DESC },
        reply_all: { type: 'boolean', description: 'Reply to all recipients (default: false)' },
        attachments: { type: 'array', description: ATTACHMENTS_DESC },
      },
      required: ['id', 'body'],
    },
  },
  {
    name: 'forward_email',
    description: 'Forward an email to new recipients. The original message and its attachments are included automatically; you can add an optional comment above and additional attachments.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'The email ID to forward' },
        to: { type: 'string', description: 'Recipient email addresses, comma-separated' },
        body: { type: 'string', description: 'Comment to add above the forwarded message (optional)' },
        body_type: { type: 'string', description: BODY_TYPE_DESC },
        cc: { type: 'string', description: 'CC recipients, comma-separated (optional)' },
        bcc: { type: 'string', description: 'BCC recipients, comma-separated (optional)' },
        attachments: { type: 'array', description: 'Additional attachments on top of the original. ' + ATTACHMENTS_DESC },
      },
      required: ['id', 'to'],
    },
  },
  {
    name: 'create_draft',
    description: 'Create a new email draft (saved but not sent). Returns the draft ID for later editing or sending.',
    inputSchema: {
      type: 'object',
      properties: {
        to: { type: 'string', description: 'Recipient email address. Separate multiple with commas.' },
        subject: { type: 'string', description: 'Email subject' },
        body: { type: 'string', description: 'Email body' },
        body_type: { type: 'string', description: BODY_TYPE_DESC },
        cc: { type: 'string', description: 'CC recipients, comma-separated (optional)' },
        bcc: { type: 'string', description: 'BCC recipients, comma-separated (optional)' },
        attachments: { type: 'array', description: ATTACHMENTS_DESC },
      },
      required: ['subject', 'body'],
    },
  },
  {
    name: 'update_draft',
    description: 'Update an existing draft email. Only provide the fields you want to change. To replace attachments, provide the full new attachment list (existing attachments will be removed first).',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'The draft email ID' },
        to: { type: 'string', description: 'New recipient list, comma-separated (replaces existing)' },
        subject: { type: 'string', description: 'New subject' },
        body: { type: 'string', description: 'New body' },
        body_type: { type: 'string', description: BODY_TYPE_DESC },
        cc: { type: 'string', description: 'New CC list, comma-separated (replaces existing)' },
        bcc: { type: 'string', description: 'New BCC list, comma-separated (replaces existing)' },
        attachments: { type: 'array', description: 'If provided, replaces all attachments. ' + ATTACHMENTS_DESC },
      },
      required: ['id'],
    },
  },
  {
    name: 'send_draft',
    description: 'Send an existing draft email by its ID.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'The draft email ID to send' },
      },
      required: ['id'],
    },
  },
  {
    name: 'schedule_send',
    description: 'Send an email at a future date/time. The Microsoft 365 mail server holds the message until the scheduled time. Same options as send_email plus a send_at datetime.',
    inputSchema: {
      type: 'object',
      properties: {
        to: { type: 'string', description: 'Recipient email address. Separate multiple with commas.' },
        subject: { type: 'string', description: 'Email subject' },
        body: { type: 'string', description: 'Email body' },
        body_type: { type: 'string', description: BODY_TYPE_DESC },
        cc: { type: 'string', description: 'CC recipients, comma-separated (optional)' },
        bcc: { type: 'string', description: 'BCC recipients, comma-separated (optional)' },
        attachments: { type: 'array', description: ATTACHMENTS_DESC },
        send_at: { type: 'string', description: 'When to send. ISO 8601 datetime in UTC, e.g. 2026-05-15T09:00:00Z. Must be in the future.' },
      },
      required: ['to', 'subject', 'body', 'send_at'],
    },
  },
  {
    name: 'get_conversation',
    description: 'Get all messages in an email thread (conversation), oldest first. Returns messages from across all folders (inbox, sent, archive, etc.). Pass either a conversation_id or a message_id from any email in the thread.',
    inputSchema: {
      type: 'object',
      properties: {
        conversation_id: { type: 'string', description: 'The conversationId (preferred — fetched from list_emails or read_email)' },
        message_id: { type: 'string', description: 'Alternatively, the ID of any single message in the thread; the conversationId will be looked up from it' },
        count: { type: 'number', description: 'Max messages to return (default: 50)' },
      },
      required: [],
    },
  },
  {
    name: 'delete_email',
    description: 'Delete an email (moves it to the Deleted Items folder).',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'The email ID to delete' },
      },
      required: ['id'],
    },
  },
  {
    name: 'move_email',
    description: 'Move an email to a different folder.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'The email ID to move' },
        destination_folder: { type: 'string', description: 'Destination: inbox, sentitems, drafts, deleteditems, or a folder ID' },
      },
      required: ['id', 'destination_folder'],
    },
  },
  // ── Calendar ───────────────────────────────────────────────────────────────
  {
    name: 'list_calendar_events',
    description: 'Retrieve calendar events within a date range.',
    inputSchema: {
      type: 'object',
      properties: {
        start_date: { type: 'string', description: 'Start of range in ISO 8601 (default: today)' },
        end_date: { type: 'string', description: 'End of range in ISO 8601 (default: 7 days from now)' },
        count: { type: 'number', description: 'Max events to return (default: 20)' },
      },
      required: [],
    },
  },
  {
    name: 'create_calendar_event',
    description: 'Create a new calendar event. Supports rich-text descriptions (HTML), recurring schedules, categories, and Teams meetings.',
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string', description: 'Event title' },
        start_datetime: { type: 'string', description: 'Start date/time in ISO 8601 (e.g. 2025-06-01T14:00:00)' },
        end_datetime: { type: 'string', description: 'End date/time in ISO 8601' },
        timezone: { type: 'string', description: 'Timezone name (e.g. Pacific/Auckland). Default: UTC' },
        location: { type: 'string', description: 'Event location or meeting room' },
        description: { type: 'string', description: 'Event description/body' },
        body_type: { type: 'string', description: BODY_TYPE_DESC },
        attendees: { type: 'string', description: 'Comma-separated attendee email addresses' },
        is_online_meeting: { type: 'boolean', description: 'Create as a Teams meeting (default: false)' },
        categories: { type: 'array', description: 'Outlook category labels (array of strings, e.g. ["Work", "Personal"])' },
        recurrence: { type: 'object', description: 'Recurrence pattern. Shape: {pattern: "daily"|"weekly"|"monthly"|"yearly", interval?, days_of_week?: ["monday"...] (weekly only), day_of_month?: 1-31 (monthly/yearly), month?: 1-12 (yearly only), end_date?: "YYYY-MM-DD" OR occurrences?: number}. Omit for one-off event.' },
      },
      required: ['subject', 'start_datetime', 'end_datetime'],
    },
  },
  {
    name: 'update_calendar_event',
    description: 'Update an existing calendar event. Only provide the fields you want to change. To remove an attendee, provide the full updated attendee list without them. To edit a single occurrence of a recurring series, pass the occurrence ID (use list_event_occurrences to find it). To edit the whole series, pass the seriesMasterId.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'The event ID. For recurring series: pass occurrence ID for single-instance edit, seriesMasterId for whole-series edit.' },
        subject: { type: 'string', description: 'New event title' },
        start_datetime: { type: 'string', description: 'New start date/time in ISO 8601' },
        end_datetime: { type: 'string', description: 'New end date/time in ISO 8601' },
        timezone: { type: 'string', description: 'Timezone name (e.g. Pacific/Auckland)' },
        location: { type: 'string', description: 'New location' },
        description: { type: 'string', description: 'New description' },
        body_type: { type: 'string', description: BODY_TYPE_DESC },
        attendees: { type: 'string', description: 'Full updated comma-separated attendee list (replaces all existing attendees)' },
        categories: { type: 'array', description: 'Outlook category labels (replaces existing categories)' },
      },
      required: ['id'],
    },
  },
  {
    name: 'delete_calendar_event',
    description: 'Delete a calendar event by its ID. Removes the event silently without sending notifications. To formally cancel a meeting and notify attendees, use cancel_calendar_event instead.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'The event ID to delete' },
      },
      required: ['id'],
    },
  },
  {
    name: 'cancel_calendar_event',
    description: 'Formally cancel a meeting you organised. Sends cancellation notices to all attendees with an optional comment. Only works for events you are the organizer of — to remove an event you do not organise, use delete_calendar_event.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'The event ID to cancel' },
        comment: { type: 'string', description: 'Cancellation message included in the notification (optional)' },
      },
      required: ['id'],
    },
  },
  {
    name: 'respond_to_event',
    description: 'Accept, decline, or tentatively accept a meeting invite. The event must be one you have been invited to (you are an attendee, not the organiser).',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'The event ID' },
        response: { type: 'string', description: '"accept", "decline", or "tentatively_accept"' },
        comment: { type: 'string', description: 'Optional comment included with the response' },
        send_response: { type: 'boolean', description: 'Whether to email a response to the organiser (default: true)' },
      },
      required: ['id', 'response'],
    },
  },
  {
    name: 'list_event_occurrences',
    description: 'List the individual occurrences of a recurring event series within a date range. Use this to find the occurrence ID for editing or cancelling a single instance of a recurring event.',
    inputSchema: {
      type: 'object',
      properties: {
        series_id: { type: 'string', description: 'The seriesMasterId of the recurring event (returned by list_calendar_events for recurring events)' },
        start_date: { type: 'string', description: 'Start of range in ISO 8601 (e.g. 2026-05-01T00:00:00Z)' },
        end_date: { type: 'string', description: 'End of range in ISO 8601' },
        count: { type: 'number', description: 'Max occurrences to return (default: 50)' },
      },
      required: ['series_id', 'start_date', 'end_date'],
    },
  },
  // ── Contacts ───────────────────────────────────────────────────────────────
  {
    name: 'list_contacts',
    description: 'List contacts from your Outlook address book, with optional search.',
    inputSchema: {
      type: 'object',
      properties: {
        search: { type: 'string', description: 'Search by name or email address' },
        count: { type: 'number', description: 'Number of contacts to return (default: 20)' },
      },
      required: [],
    },
  },
  {
    name: 'create_contact',
    description: 'Create a new contact in Outlook.',
    inputSchema: {
      type: 'object',
      properties: {
        first_name: { type: 'string', description: 'First name' },
        last_name: { type: 'string', description: 'Last name' },
        email: { type: 'string', description: 'Email address' },
        phone: { type: 'string', description: 'Phone number (optional)' },
        company: { type: 'string', description: 'Company name (optional)' },
        job_title: { type: 'string', description: 'Job title (optional)' },
      },
      required: ['first_name', 'last_name'],
    },
  },
  // ── Tasks ──────────────────────────────────────────────────────────────────
  {
    name: 'list_task_lists',
    description: 'List all Microsoft To Do task lists. Use this to find list IDs before listing or creating tasks.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
  },
  {
    name: 'list_tasks',
    description: 'List tasks in a To Do task list. Defaults to the default task list and excludes completed tasks.',
    inputSchema: {
      type: 'object',
      properties: {
        list_id: { type: 'string', description: 'Task list ID (use list_task_lists to find IDs). Defaults to default list.' },
        status: { type: 'string', description: 'Filter by status: notStarted, inProgress, completed. Default: excludes completed.' },
      },
      required: [],
    },
  },
  {
    name: 'create_task',
    description: 'Create a new task in a Microsoft To Do list.',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Task title' },
        list_id: { type: 'string', description: 'Task list ID. Defaults to default task list.' },
        due_date: { type: 'string', description: 'Due date in ISO 8601 (optional)' },
        body: { type: 'string', description: 'Task notes or description (optional)' },
        importance: { type: 'string', description: 'Priority: low, normal, high (default: normal)' },
      },
      required: ['title'],
    },
  },
  // ── Mailbox settings ───────────────────────────────────────────────────────
  {
    name: 'get_mailbox_settings',
    description: 'Get mailbox settings including out-of-office status, timezone, and working hours.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
  },
  {
    name: 'set_out_of_office',
    description: 'Enable or disable the automatic out-of-office reply.',
    inputSchema: {
      type: 'object',
      properties: {
        enabled: { type: 'boolean', description: 'true to enable, false to disable' },
        internal_message: { type: 'string', description: 'Auto-reply for people inside your organisation' },
        external_message: { type: 'string', description: 'Auto-reply for people outside your organisation (defaults to internal message)' },
        start_datetime: { type: 'string', description: 'Scheduled start in ISO 8601 (optional — omit to activate immediately)' },
        end_datetime: { type: 'string', description: 'Scheduled end in ISO 8601 (optional)' },
      },
      required: ['enabled'],
    },
  },
  // ── OneDrive ───────────────────────────────────────────────────────────────
  {
    name: 'list_files',
    description: 'List files and folders in OneDrive.',
    inputSchema: {
      type: 'object',
      properties: {
        folder_path: { type: 'string', description: 'Path to a subfolder (e.g. "Documents/Projects"). Defaults to root.' },
        count: { type: 'number', description: 'Number of items to return (default: 25)' },
      },
      required: [],
    },
  },
  {
    name: 'share_file',
    description: 'Create a shareable link for a file or folder in OneDrive.',
    inputSchema: {
      type: 'object',
      properties: {
        item_path: { type: 'string', description: 'Path to the file or folder (e.g. "Documents/report.pdf")' },
        link_type: { type: 'string', description: 'view (read-only) or edit (read-write). Default: view' },
        scope: { type: 'string', description: 'anonymous (anyone with the link) or organization (org members only). Default: anonymous' },
      },
      required: ['item_path'],
    },
  },
];

// ── Tool dispatcher ────────────────────────────────────────────────────────────

export async function callTool(name: string, args: ToolArgs, env: Env): Promise<string> {
  const token = await getValidAccessToken(env);

  switch (name) {
    case 'list_emails':            return listEmails(token, args);
    case 'read_email':             return readEmail(token, args);
    case 'send_email':             return sendEmail(token, args);
    case 'reply_to_email':         return replyToEmail(token, args);
    case 'forward_email':          return forwardEmail(token, args);
    case 'delete_email':           return deleteEmail(token, args);
    case 'move_email':             return moveEmail(token, args);
    case 'create_draft':           return createDraft(token, args);
    case 'update_draft':           return updateDraft(token, args);
    case 'send_draft':             return sendDraft(token, args);
    case 'schedule_send':          return scheduleSend(token, args);
    case 'get_conversation':       return getConversation(token, args);
    case 'list_calendar_events':   return listCalendarEvents(token, args);
    case 'create_calendar_event':  return createCalendarEvent(token, args);
    case 'update_calendar_event':  return updateCalendarEvent(token, args);
    case 'delete_calendar_event':  return deleteCalendarEvent(token, args);
    case 'cancel_calendar_event':  return cancelCalendarEvent(token, args);
    case 'respond_to_event':       return respondToEvent(token, args);
    case 'list_event_occurrences': return listEventOccurrences(token, args);
    case 'list_contacts':          return listContacts(token, args);
    case 'create_contact':         return createContact(token, args);
    case 'list_task_lists':        return listTaskLists(token);
    case 'list_tasks':             return listTasks(token, args);
    case 'create_task':            return createTask(token, args);
    case 'get_mailbox_settings':   return getMailboxSettings(token);
    case 'set_out_of_office':      return setOutOfOffice(token, args);
    case 'list_files':             return listFiles(token, args);
    case 'share_file':             return shareFile(token, args);
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

// ── Email handlers ─────────────────────────────────────────────────────────────

async function listEmails(token: string, args: ToolArgs): Promise<string> {
  const count = Math.min((args.count as number) ?? 20, 50);
  const folder = (args.folder as string) ?? 'inbox';
  const search = args.search as string | undefined;

  let path: string;
  if (search) {
    path = `/me/messages?$search="${search}"&$select=id,conversationId,subject,from,receivedDateTime,bodyPreview,isRead,isDraft&$top=${count}`;
  } else {
    path = `/me/mailFolders/${folder}/messages?$select=id,conversationId,subject,from,receivedDateTime,bodyPreview,isRead,isDraft&$top=${count}&$orderby=receivedDateTime desc`;
  }

  const data = await graphRequest(token, path) as { value: unknown[] };
  return JSON.stringify(sanitizeEmailList(data.value));
}

async function readEmail(token: string, args: ToolArgs): Promise<string> {
  const id = args.id as string;
  const data = await graphRequest(
    token,
    `/me/messages/${id}?$select=id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,isDraft`
  );
  return JSON.stringify(sanitizeEmailFull(data));
}

// Build the Graph "message" object from common args. Attachments are NOT
// included here — they are uploaded separately via uploadAttachmentsToDraft
// because they may need server-side fetching (OneDrive / URL).
function buildMessageObject(args: ToolArgs): Record<string, unknown> {
  const message: Record<string, unknown> = {
    subject: args.subject,
    body: buildMessageBody(args.body as string, args.body_type as string | undefined),
    toRecipients: buildRecipients(args.to as string | undefined),
  };
  const cc = buildRecipients(args.cc as string | undefined);
  const bcc = buildRecipients(args.bcc as string | undefined);
  if (cc.length) message.ccRecipients = cc;
  if (bcc.length) message.bccRecipients = bcc;
  return message;
}

async function sendEmail(token: string, args: ToolArgs): Promise<string> {
  // If the message has attachments we must use the create-draft-then-send
  // pattern, because /me/sendMail enforces a 4 MB total request cap. Drafts
  // can have attachments uploaded individually and are not subject to that
  // single-request cap.
  const hasAttachments = Array.isArray(args.attachments) && args.attachments.length > 0;

  if (!hasAttachments) {
    await graphRequest(token, '/me/sendMail', {
      method: 'POST',
      body: JSON.stringify({ message: buildMessageObject(args) }),
    });
    auditLog('email_sent', { to: args.to, subject: args.subject, attachments: 0 });
    return JSON.stringify({ success: true, message: 'Email sent.' });
  }

  // Path with attachments: create draft (without attachments), upload each
  // attachment, then send the draft.
  const argsNoAttachments = { ...args, attachments: undefined };
  const draft = await graphRequest(token, '/me/messages', {
    method: 'POST',
    body: JSON.stringify(buildMessageObject(argsNoAttachments)),
  }) as { id: string };

  await uploadAttachmentsToDraft(token, draft.id, args.attachments);
  await graphRequest(token, `/me/messages/${draft.id}/send`, { method: 'POST' });
  auditLog('email_sent', {
    to: args.to,
    subject: args.subject,
    attachments: (args.attachments as unknown[]).length,
    attachment_sources: summariseAttachmentSources(args.attachments),
  });
  return JSON.stringify({ success: true, message: 'Email sent.' });
}

// Structured audit logging for every outbound-mail path. Captured by
// Cloudflare's observability (3-day retention) plus visible via wrangler tail.
function auditLog(event: string, details: Record<string, unknown>): void {
  console.log(JSON.stringify({ audit: event, ts: new Date().toISOString(), ...details }));
}

// Summarise attachment sources without leaking full paths/URLs.
// Returns e.g. {onedrive: 2, url: 1, base64: 0}.
function summariseAttachmentSources(raw: unknown): Record<string, number> {
  if (!Array.isArray(raw)) return {};
  const out = { onedrive: 0, url: 0, base64: 0 };
  for (const a of raw as Array<Record<string, unknown>>) {
    if (a.onedrive_path) out.onedrive++;
    else if (a.url) out.url++;
    else if (a.content_base64) out.base64++;
  }
  return out;
}

async function replyToEmail(token: string, args: ToolArgs): Promise<string> {
  const id = args.id as string;
  const replyAll = (args.reply_all as boolean) ?? false;
  const action = replyAll ? 'createReplyAll' : 'createReply';

  // Step 1: create a reply draft (Graph pre-fills recipients and quoted body)
  const draft = await graphRequest(token, `/me/messages/${id}/${action}`, {
    method: 'POST',
  }) as { id: string };

  // Step 2: set the body. Use PATCH to update only the body — Graph preserves
  // the auto-generated quoted history below.
  const body = buildMessageBody(args.body as string, args.body_type as string | undefined);
  await graphRequest(token, `/me/messages/${draft.id}`, {
    method: 'PATCH',
    body: JSON.stringify({ body }),
  });

  // Step 3: upload attachments if any
  if (args.attachments) {
    await uploadAttachmentsToDraft(token, draft.id, args.attachments);
  }

  // Step 4: send
  await graphRequest(token, `/me/messages/${draft.id}/send`, { method: 'POST' });
  auditLog('reply_sent', {
    in_reply_to: id,
    reply_all: replyAll,
    attachments: Array.isArray(args.attachments) ? args.attachments.length : 0,
    attachment_sources: summariseAttachmentSources(args.attachments),
  });
  return JSON.stringify({ success: true, message: 'Reply sent.' });
}

async function forwardEmail(token: string, args: ToolArgs): Promise<string> {
  const id = args.id as string;
  if (!args.to) throw new Error('"to" is required for forward_email');

  // Step 1: create a forward draft (Graph copies the original message + its attachments)
  const draft = await graphRequest(token, `/me/messages/${id}/createForward`, {
    method: 'POST',
  }) as { id: string };

  // Step 2: set recipients and (optional) comment body via PATCH.
  // PATCHing body REPLACES the whole body — Graph then re-appends the original
  // forwarded message below automatically when /send runs.
  const updates: Record<string, unknown> = {
    toRecipients: buildRecipients(args.to as string),
  };
  const cc = buildRecipients(args.cc as string | undefined);
  const bcc = buildRecipients(args.bcc as string | undefined);
  if (cc.length) updates.ccRecipients = cc;
  if (bcc.length) updates.bccRecipients = bcc;
  if (args.body) {
    updates.body = buildMessageBody(args.body as string, args.body_type as string | undefined);
  }
  await graphRequest(token, `/me/messages/${draft.id}`, {
    method: 'PATCH',
    body: JSON.stringify(updates),
  });

  // Step 3: upload any additional attachments
  if (args.attachments) {
    await uploadAttachmentsToDraft(token, draft.id, args.attachments);
  }

  // Step 4: send
  await graphRequest(token, `/me/messages/${draft.id}/send`, { method: 'POST' });
  auditLog('email_forwarded', {
    forwarded_from: id,
    to: args.to,
    additional_attachments: Array.isArray(args.attachments) ? args.attachments.length : 0,
    attachment_sources: summariseAttachmentSources(args.attachments),
  });
  return JSON.stringify({ success: true, message: 'Email forwarded.' });
}

async function uploadAttachmentsToDraft(
  token: string,
  draftId: string,
  attachmentsRaw: unknown
): Promise<void> {
  const attachments = await resolveAndValidateAttachments(token, attachmentsRaw);
  for (const att of attachments) {
    await graphRequest(token, `/me/messages/${draftId}/attachments`, {
      method: 'POST',
      body: JSON.stringify(att),
    });
  }
}

// ── Draft handlers ─────────────────────────────────────────────────────────────

async function createDraft(token: string, args: ToolArgs): Promise<string> {
  const argsNoAttachments = { ...args, attachments: undefined };
  const draft = await graphRequest(token, '/me/messages', {
    method: 'POST',
    body: JSON.stringify(buildMessageObject(argsNoAttachments)),
  }) as { id: string; webLink?: string };

  if (args.attachments) {
    await uploadAttachmentsToDraft(token, draft.id, args.attachments);
  }

  auditLog('draft_created', { draft_id: draft.id, to: args.to, subject: args.subject, attachment_sources: summariseAttachmentSources(args.attachments) });
  return JSON.stringify({ success: true, message: 'Draft created.', draft_id: draft.id, webLink: draft.webLink });
}

async function updateDraft(token: string, args: ToolArgs): Promise<string> {
  const id = args.id as string;
  const updates: Record<string, unknown> = {};

  if (args.subject) updates.subject = args.subject;
  if (args.body) updates.body = buildMessageBody(args.body as string, args.body_type as string | undefined);
  if (args.to) updates.toRecipients = buildRecipients(args.to as string);
  if (args.cc) updates.ccRecipients = buildRecipients(args.cc as string);
  if (args.bcc) updates.bccRecipients = buildRecipients(args.bcc as string);

  if (Object.keys(updates).length > 0) {
    await graphRequest(token, `/me/messages/${id}`, {
      method: 'PATCH',
      body: JSON.stringify(updates),
    });
  }

  // Replace attachments if provided. Graph requires deleting then re-uploading
  // because attachments are not directly patchable.
  if (args.attachments !== undefined) {
    const existing = await graphRequest(token, `/me/messages/${id}/attachments?$select=id`) as { value: Array<{ id: string }> };
    for (const a of existing.value) {
      await graphRequest(token, `/me/messages/${id}/attachments/${a.id}`, { method: 'DELETE' });
    }
    if (args.attachments) {
      await uploadAttachmentsToDraft(token, id, args.attachments);
    }
  }

  auditLog('draft_updated', { draft_id: id });
  return JSON.stringify({ success: true, message: 'Draft updated.' });
}

async function sendDraft(token: string, args: ToolArgs): Promise<string> {
  const id = args.id as string;
  await graphRequest(token, `/me/messages/${id}/send`, { method: 'POST' });
  auditLog('draft_sent', { draft_id: id });
  return JSON.stringify({ success: true, message: 'Draft sent.' });
}

// ── Scheduled send ─────────────────────────────────────────────────────────────
// Graph supports deferred send via the MAPI extended property
// PR_DEFERRED_SEND_TIME (PtagId 0x3FEFSystemTime). We set it on a draft, then
// /send queues the message — the M365 server holds it until the deferred time.

async function scheduleSend(token: string, args: ToolArgs): Promise<string> {
  const sendAt = args.send_at as string;
  if (!sendAt) throw new Error('send_at is required (ISO 8601 UTC datetime)');
  const sendAtDate = new Date(sendAt);
  if (isNaN(sendAtDate.getTime())) throw new Error('send_at is not a valid ISO 8601 datetime');
  if (sendAtDate.getTime() <= Date.now()) throw new Error('send_at must be in the future');

  // Create the draft with the deferred-send extended property
  const argsNoAttachments = { ...args, attachments: undefined };
  const message = buildMessageObject(argsNoAttachments);
  message.singleValueExtendedProperties = [
    {
      id: 'SystemTime 0x3FEF',
      value: sendAtDate.toISOString(),
    },
  ];

  const draft = await graphRequest(token, '/me/messages', {
    method: 'POST',
    body: JSON.stringify(message),
  }) as { id: string };

  if (args.attachments) {
    await uploadAttachmentsToDraft(token, draft.id, args.attachments);
  }

  await graphRequest(token, `/me/messages/${draft.id}/send`, { method: 'POST' });
  auditLog('scheduled_send', {
    draft_id: draft.id,
    send_at: sendAtDate.toISOString(),
    to: args.to,
    subject: args.subject,
    attachment_sources: summariseAttachmentSources(args.attachments),
  });
  return JSON.stringify({
    success: true,
    message: `Email scheduled for ${sendAtDate.toISOString()}. The mail server will hold it until then.`,
    scheduled_id: draft.id,
  });
}

// ── Conversation / threading ───────────────────────────────────────────────────

async function getConversation(token: string, args: ToolArgs): Promise<string> {
  let conversationId = args.conversation_id as string | undefined;

  // If only a message_id was given, look up its conversationId first
  if (!conversationId) {
    const messageId = args.message_id as string | undefined;
    if (!messageId) throw new Error('Either conversation_id or message_id is required');
    const msg = await graphRequest(token, `/me/messages/${messageId}?$select=conversationId`) as { conversationId?: string };
    if (!msg.conversationId) throw new Error('Could not resolve conversationId from message');
    conversationId = msg.conversationId;
  }

  const count = (args.count as number) ?? 50;
  // Escape single quotes in the ID for the OData filter, then URL-encode
  // the whole filter value (conversationIds contain `=` and `+`).
  const safeId = conversationId.replace(/'/g, "''");
  const filter = encodeURIComponent(`conversationId eq '${safeId}'`);
  // Note: NOT using $orderby — combining $filter with $orderby on a different
  // property requires the "advanced query" opt-in header and Graph still
  // rejects it for some mailbox types. Sort client-side instead.
  const path =
    `/me/messages` +
    `?$filter=${filter}` +
    `&$select=id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,isDraft` +
    `&$top=${count}`;

  const data = await graphRequest(token, path) as { value: Array<{ receivedDateTime?: string }> };

  // Sort oldest-first by receivedDateTime client-side
  const sorted = [...data.value].sort((a, b) => {
    const ta = a.receivedDateTime ? Date.parse(a.receivedDateTime) : 0;
    const tb = b.receivedDateTime ? Date.parse(b.receivedDateTime) : 0;
    return ta - tb;
  });

  return JSON.stringify({
    conversationId,
    count: sorted.length,
    messages: sorted.map(m => sanitizeEmailFull(m)),
  });
}

async function deleteEmail(token: string, args: ToolArgs): Promise<string> {
  await graphRequest(token, `/me/messages/${args.id as string}`, { method: 'DELETE' });
  return JSON.stringify({ success: true, message: 'Email deleted.' });
}

async function moveEmail(token: string, args: ToolArgs): Promise<string> {
  const data = await graphRequest(token, `/me/messages/${args.id as string}/move`, {
    method: 'POST',
    body: JSON.stringify({ destinationId: args.destination_folder }),
  }) as { id: string };
  return JSON.stringify({ success: true, message: 'Email moved.', new_id: data.id });
}

// ── Calendar handlers ──────────────────────────────────────────────────────────

// Common $select for events — kept here so list and instances stay in sync.
const EVENT_SELECT = 'id,seriesMasterId,type,subject,start,end,location,attendees,organizer,bodyPreview,isOnlineMeeting,onlineMeetingUrl,categories,recurrence,isCancelled,responseStatus';

async function listCalendarEvents(token: string, args: ToolArgs): Promise<string> {
  const now = new Date();
  const weekLater = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  const start = (args.start_date as string) ?? now.toISOString();
  const end = (args.end_date as string) ?? weekLater.toISOString();
  const count = (args.count as number) ?? 20;

  const path =
    `/me/events` +
    `?$select=${EVENT_SELECT}` +
    `&$top=${count}` +
    `&$filter=start/dateTime ge '${start}' and start/dateTime le '${end}'` +
    `&$orderby=start/dateTime`;

  const data = await graphRequest(token, path) as { value: unknown[] };
  return JSON.stringify(sanitizeEventList(data.value));
}

async function createCalendarEvent(token: string, args: ToolArgs): Promise<string> {
  const timezone = (args.timezone as string) ?? 'UTC';
  const attendeeList = args.attendees
    ? (args.attendees as string).split(',').map(e => e.trim())
    : [];

  const event: Record<string, unknown> = {
    subject: args.subject,
    start: { dateTime: args.start_datetime, timeZone: timezone },
    end: { dateTime: args.end_datetime, timeZone: timezone },
  };

  if (args.location) event.location = { displayName: args.location };
  if (args.description) {
    // Reuses the email outbound HTML sanitiser when body_type is "html"
    event.body = buildMessageBody(args.description as string, args.body_type as string | undefined);
  }
  if (attendeeList.length > 0) {
    event.attendees = attendeeList.map(email => ({
      emailAddress: { address: email },
      type: 'required',
    }));
  }
  if (args.is_online_meeting) {
    event.isOnlineMeeting = true;
    event.onlineMeetingProvider = 'teamsForBusiness';
  }
  if (Array.isArray(args.categories) && args.categories.length > 0) {
    event.categories = args.categories;
  }
  if (args.recurrence) {
    event.recurrence = buildRecurrence(
      args.recurrence as RecurrenceInput,
      args.start_datetime as string,
      timezone
    );
  }

  const data = await graphRequest(token, '/me/events', {
    method: 'POST',
    body: JSON.stringify(event),
  });

  return JSON.stringify({ success: true, message: 'Event created.', event: data });
}

async function updateCalendarEvent(token: string, args: ToolArgs): Promise<string> {
  const id = args.id as string;
  const timezone = (args.timezone as string) ?? 'UTC';
  const updates: Record<string, unknown> = {};

  if (args.subject) updates.subject = args.subject;
  if (args.start_datetime) updates.start = { dateTime: args.start_datetime, timeZone: timezone };
  if (args.end_datetime) updates.end = { dateTime: args.end_datetime, timeZone: timezone };
  if (args.location) updates.location = { displayName: args.location };
  if (args.description) {
    updates.body = buildMessageBody(args.description as string, args.body_type as string | undefined);
  }
  if (args.attendees) {
    updates.attendees = (args.attendees as string)
      .split(',')
      .map(e => ({ emailAddress: { address: e.trim() }, type: 'required' }));
  }
  if (Array.isArray(args.categories)) {
    updates.categories = args.categories;
  }

  await graphRequest(token, `/me/events/${id}`, {
    method: 'PATCH',
    body: JSON.stringify(updates),
  });

  return JSON.stringify({ success: true, message: 'Event updated.' });
}

async function deleteCalendarEvent(token: string, args: ToolArgs): Promise<string> {
  await graphRequest(token, `/me/events/${args.id as string}`, { method: 'DELETE' });
  return JSON.stringify({ success: true, message: 'Event deleted.' });
}

async function cancelCalendarEvent(token: string, args: ToolArgs): Promise<string> {
  const id = args.id as string;
  const body = args.comment ? { comment: args.comment } : {};
  await graphRequest(token, `/me/events/${id}/cancel`, {
    method: 'POST',
    body: JSON.stringify(body),
  });
  auditLog('event_cancelled', { event_id: id, has_comment: !!args.comment });
  return JSON.stringify({ success: true, message: 'Event cancelled and attendees notified.' });
}

async function respondToEvent(token: string, args: ToolArgs): Promise<string> {
  const id = args.id as string;
  const response = (args.response as string).toLowerCase();

  // Map our friendly names to Graph endpoints
  let endpoint: string;
  if (response === 'accept') endpoint = 'accept';
  else if (response === 'decline') endpoint = 'decline';
  else if (response === 'tentatively_accept' || response === 'tentative') endpoint = 'tentativelyAccept';
  else throw new Error(`Invalid response "${args.response}". Use: accept, decline, tentatively_accept`);

  const body: Record<string, unknown> = {
    sendResponse: (args.send_response as boolean) ?? true,
  };
  if (args.comment) body.comment = args.comment;

  await graphRequest(token, `/me/events/${id}/${endpoint}`, {
    method: 'POST',
    body: JSON.stringify(body),
  });

  auditLog('event_response', { event_id: id, response, send_response: body.sendResponse });
  return JSON.stringify({ success: true, message: `Responded with: ${response}` });
}

async function listEventOccurrences(token: string, args: ToolArgs): Promise<string> {
  const seriesId = args.series_id as string;
  const startDate = args.start_date as string;
  const endDate = args.end_date as string;
  const count = (args.count as number) ?? 50;

  const path =
    `/me/events/${seriesId}/instances` +
    `?startDateTime=${encodeURIComponent(startDate)}` +
    `&endDateTime=${encodeURIComponent(endDate)}` +
    `&$select=${EVENT_SELECT}` +
    `&$top=${count}` +
    `&$orderby=start/dateTime`;

  const data = await graphRequest(token, path) as { value: unknown[] };
  return JSON.stringify({
    series_id: seriesId,
    count: data.value.length,
    occurrences: sanitizeEventList(data.value),
  });
}

// ── Contact handlers ───────────────────────────────────────────────────────────

async function listContacts(token: string, args: ToolArgs): Promise<string> {
  const count = (args.count as number) ?? 20;
  let path = `/me/contacts?$select=id,displayName,emailAddresses,phones,companyName,jobTitle&$top=${count}&$orderby=displayName`;

  if (args.search) {
    const q = args.search as string;
    path += `&$filter=contains(displayName,'${q}')`;
  }

  const data = await graphRequest(token, path) as { value: unknown[] };
  return JSON.stringify(sanitizeContactList(data.value));
}

async function createContact(token: string, args: ToolArgs): Promise<string> {
  const contact: Record<string, unknown> = {
    givenName: args.first_name,
    surname: args.last_name,
  };

  if (args.email) {
    contact.emailAddresses = [{
      address: args.email,
      name: `${args.first_name} ${args.last_name}`,
    }];
  }
  if (args.phone) contact.phones = [{ number: args.phone, type: 'mobile' }];
  if (args.company) contact.companyName = args.company;
  if (args.job_title) contact.jobTitle = args.job_title;

  const data = await graphRequest(token, '/me/contacts', {
    method: 'POST',
    body: JSON.stringify(contact),
  });

  return JSON.stringify({ success: true, message: 'Contact created.', contact: data });
}

// ── Task handlers ──────────────────────────────────────────────────────────────

async function listTaskLists(token: string): Promise<string> {
  const data = await graphRequest(
    token,
    '/me/todo/lists?$select=id,displayName,isOwner,isShared,wellknownListName'
  ) as { value: unknown[] };
  return JSON.stringify(sanitizeTaskLists(data.value));
}

async function resolveTaskListId(token: string, listId?: string): Promise<string> {
  if (listId) return listId;
  const data = await graphRequest(token, '/me/todo/lists') as {
    value: Array<{ id: string; wellknownListName: string }>;
  };
  return data.value.find(l => l.wellknownListName === 'defaultList')?.id ?? data.value[0].id;
}

async function listTasks(token: string, args: ToolArgs): Promise<string> {
  const listId = await resolveTaskListId(token, args.list_id as string | undefined);
  let path = `/me/todo/lists/${listId}/tasks?$select=id,title,status,importance,dueDateTime,body,createdDateTime`;

  if (args.status) {
    path += `&$filter=status eq '${args.status}'`;
  } else {
    path += `&$filter=status ne 'completed'`;
  }

  const data = await graphRequest(token, path) as { value: unknown[] };
  return JSON.stringify(sanitizeTaskList(data.value));
}

async function createTask(token: string, args: ToolArgs): Promise<string> {
  const listId = await resolveTaskListId(token, args.list_id as string | undefined);

  const task: Record<string, unknown> = {
    title: args.title,
    importance: (args.importance as string) ?? 'normal',
  };

  if (args.due_date) task.dueDateTime = { dateTime: args.due_date, timeZone: 'UTC' };
  if (args.body) task.body = { content: args.body, contentType: 'text' };

  const data = await graphRequest(token, `/me/todo/lists/${listId}/tasks`, {
    method: 'POST',
    body: JSON.stringify(task),
  });

  return JSON.stringify({ success: true, message: 'Task created.', task: data });
}

// ── Mailbox settings handlers ──────────────────────────────────────────────────

async function getMailboxSettings(token: string): Promise<string> {
  const data = await graphRequest(token, '/me/mailboxSettings');
  return JSON.stringify(sanitizeMailboxSettings(data));
}

async function setOutOfOffice(token: string, args: ToolArgs): Promise<string> {
  const enabled = args.enabled as boolean;
  const internalMsg = args.internal_message as string | undefined;
  const externalMsg = (args.external_message as string | undefined) ?? internalMsg;

  const replySettings: Record<string, unknown> = {
    status: enabled ? 'alwaysEnabled' : 'disabled',
  };

  if (enabled && internalMsg) {
    replySettings.internalReplyMessage = internalMsg;
    replySettings.externalReplyMessage = externalMsg;
  }

  if (enabled && args.start_datetime && args.end_datetime) {
    replySettings.status = 'scheduled';
    replySettings.scheduledStartDateTime = { dateTime: args.start_datetime, timeZone: 'UTC' };
    replySettings.scheduledEndDateTime = { dateTime: args.end_datetime, timeZone: 'UTC' };
  }

  await graphRequest(token, '/me/mailboxSettings', {
    method: 'PATCH',
    body: JSON.stringify({ automaticRepliesSetting: replySettings }),
  });

  return JSON.stringify({ success: true, message: `Out-of-office ${enabled ? 'enabled' : 'disabled'}.` });
}

// ── OneDrive handlers ──────────────────────────────────────────────────────────

async function listFiles(token: string, args: ToolArgs): Promise<string> {
  const count = (args.count as number) ?? 25;
  const folderPath = args.folder_path as string | undefined;

  const basePath = folderPath
    ? `/me/drive/root:/${folderPath}:/children`
    : `/me/drive/root/children`;

  const path = `${basePath}?$select=id,name,size,file,folder,lastModifiedDateTime,webUrl&$top=${count}`;
  const data = await graphRequest(token, path) as { value: unknown[] };
  return JSON.stringify(sanitizeFileList(data.value));
}

async function shareFile(token: string, args: ToolArgs): Promise<string> {
  const itemPath = args.item_path as string;
  const linkType = (args.link_type as string) ?? 'view';
  const scope = (args.scope as string) ?? 'anonymous';

  const data = await graphRequest(
    token,
    `/me/drive/root:/${itemPath}:/createLink`,
    {
      method: 'POST',
      body: JSON.stringify({ type: linkType, scope }),
    }
  ) as { link: { webUrl: string; type: string; scope: string } };

  return JSON.stringify({
    success: true,
    url: data.link.webUrl,
    type: data.link.type,
    scope: data.link.scope,
  });
}
