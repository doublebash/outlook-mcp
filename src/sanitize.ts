// ── HTML → plain text ──────────────────────────────────────────────────────────

export function htmlToText(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<\/?(p|div|br|h[1-6]|li|tr|blockquote|hr)[^>]*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&[a-z]+;/gi, '')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// ── Helpers ────────────────────────────────────────────────────────────────────

function formatAddress(addr: { name?: string; address?: string } | undefined): string {
  if (!addr) return '';
  return addr.name ? `${addr.name} <${addr.address}>` : (addr.address ?? '');
}

function formatAddressList(
  recipients: Array<{ emailAddress?: { name?: string; address?: string } }> | undefined
): string[] {
  if (!recipients?.length) return [];
  return recipients.map(r => formatAddress(r.emailAddress));
}

// ── Email list (from list_emails) ──────────────────────────────────────────────

interface RawEmailSummary {
  id?: string;
  conversationId?: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  receivedDateTime?: string;
  bodyPreview?: string;
  isRead?: boolean;
  isDraft?: boolean;
}

export function sanitizeEmailList(items: unknown[]): unknown[] {
  return (items as RawEmailSummary[]).map(item => ({
    id: item.id,
    conversationId: item.conversationId,
    subject: item.subject ?? '(no subject)',
    from: formatAddress(item.from?.emailAddress),
    received: item.receivedDateTime,
    preview: item.bodyPreview,
    isRead: item.isRead ?? false,
    isDraft: item.isDraft ?? false,
  }));
}

// ── Email full (from read_email) ───────────────────────────────────────────────

interface RawEmailFull {
  id?: string;
  conversationId?: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  toRecipients?: Array<{ emailAddress?: { name?: string; address?: string } }>;
  ccRecipients?: Array<{ emailAddress?: { name?: string; address?: string } }>;
  receivedDateTime?: string;
  body?: { contentType?: string; content?: string };
  hasAttachments?: boolean;
  isDraft?: boolean;
}

export function sanitizeEmailFull(item: unknown): unknown {
  const e = item as RawEmailFull;
  const bodyContent = e.body?.content ?? '';
  const bodyText =
    e.body?.contentType?.toLowerCase() === 'html'
      ? htmlToText(bodyContent)
      : bodyContent;

  return {
    id: e.id,
    conversationId: e.conversationId,
    subject: e.subject ?? '(no subject)',
    from: formatAddress(e.from?.emailAddress),
    to: formatAddressList(e.toRecipients),
    cc: formatAddressList(e.ccRecipients),
    received: e.receivedDateTime,
    body: bodyText,
    hasAttachments: e.hasAttachments ?? false,
    isDraft: e.isDraft ?? false,
  };
}

// ── Calendar event list ────────────────────────────────────────────────────────

interface RawEvent {
  id?: string;
  seriesMasterId?: string;
  type?: string; // singleInstance | occurrence | exception | seriesMaster
  subject?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  location?: { displayName?: string };
  attendees?: Array<{
    emailAddress?: { name?: string; address?: string };
    type?: string;
    status?: { response?: string; time?: string };
  }>;
  organizer?: { emailAddress?: { name?: string; address?: string } };
  bodyPreview?: string;
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
  categories?: string[];
  recurrence?: unknown;
  isCancelled?: boolean;
  responseStatus?: { response?: string };
}

export function sanitizeEventList(items: unknown[]): unknown[] {
  return (items as RawEvent[]).map(item => {
    const attendees = (item.attendees ?? []).map(a => {
      const addr = formatAddress(a.emailAddress);
      const role = a.type ?? 'required';
      const status = a.status?.response ?? 'none';
      return `${addr} [${role}, ${status}]`;
    });

    return {
      id: item.id,
      seriesMasterId: item.seriesMasterId ?? undefined,
      type: item.type ?? 'singleInstance',
      subject: item.subject,
      start: item.start?.dateTime,
      end: item.end?.dateTime,
      timezone: item.start?.timeZone,
      location: item.location?.displayName ?? null,
      organizer: formatAddress(item.organizer?.emailAddress) || undefined,
      attendees: attendees.length > 0 ? attendees : undefined,
      categories: item.categories?.length ? item.categories : undefined,
      preview: item.bodyPreview,
      isOnlineMeeting: item.isOnlineMeeting ?? false,
      meetingUrl: item.onlineMeetingUrl ?? null,
      isRecurring: !!item.recurrence,
      isCancelled: item.isCancelled ?? false,
      myResponse: item.responseStatus?.response ?? undefined,
    };
  });
}

// ── Contact list ───────────────────────────────────────────────────────────────

interface RawContact {
  id?: string;
  displayName?: string;
  emailAddresses?: Array<{ address?: string }>;
  phones?: Array<{ number?: string; type?: string }>;
  companyName?: string;
  jobTitle?: string;
}

export function sanitizeContactList(items: unknown[]): unknown[] {
  return (items as RawContact[]).map(item => ({
    id: item.id,
    name: item.displayName,
    email: item.emailAddresses?.[0]?.address ?? null,
    phone: item.phones?.[0]?.number ?? null,
    company: item.companyName ?? null,
    jobTitle: item.jobTitle ?? null,
  }));
}

// ── Task list ──────────────────────────────────────────────────────────────────

interface RawTask {
  id?: string;
  title?: string;
  status?: string;
  importance?: string;
  dueDateTime?: { dateTime?: string };
  body?: { content?: string };
  createdDateTime?: string;
}

export function sanitizeTaskList(items: unknown[]): unknown[] {
  return (items as RawTask[]).map(item => ({
    id: item.id,
    title: item.title,
    status: item.status,
    importance: item.importance,
    due: item.dueDateTime?.dateTime ?? null,
    notes: item.body?.content ? htmlToText(item.body.content) : null,
    created: item.createdDateTime,
  }));
}

// ── Task lists (from list_task_lists) ─────────────────────────────────────────

interface RawTaskList {
  id?: string;
  displayName?: string;
  isOwner?: boolean;
  isShared?: boolean;
  wellknownListName?: string;
}

export function sanitizeTaskLists(items: unknown[]): unknown[] {
  return (items as RawTaskList[]).map(item => ({
    id: item.id,
    name: item.displayName,
    isDefault: item.wellknownListName === 'defaultList',
    isShared: item.isShared ?? false,
  }));
}

// ── Mailbox settings ───────────────────────────────────────────────────────────

interface RawMailboxSettings {
  timeZone?: string;
  language?: { displayName?: string; locale?: string };
  workingHours?: {
    daysOfWeek?: string[];
    startTime?: string;
    endTime?: string;
  };
  automaticRepliesSetting?: {
    status?: string;
    scheduledStartDateTime?: { dateTime?: string };
    scheduledEndDateTime?: { dateTime?: string };
    internalReplyMessage?: string;
    externalReplyMessage?: string;
  };
}

export function sanitizeMailboxSettings(item: unknown): unknown {
  const s = item as RawMailboxSettings;
  const ooo = s.automaticRepliesSetting;

  return {
    timezone: s.timeZone,
    language: s.language?.displayName,
    workingHours: s.workingHours
      ? {
          days: s.workingHours.daysOfWeek,
          start: s.workingHours.startTime,
          end: s.workingHours.endTime,
        }
      : null,
    outOfOffice: ooo
      ? {
          status: ooo.status,
          scheduledStart: ooo.scheduledStartDateTime?.dateTime ?? null,
          scheduledEnd: ooo.scheduledEndDateTime?.dateTime ?? null,
          internalMessage: ooo.internalReplyMessage
            ? htmlToText(ooo.internalReplyMessage)
            : null,
          externalMessage: ooo.externalReplyMessage
            ? htmlToText(ooo.externalReplyMessage)
            : null,
        }
      : null,
  };
}

// ── OneDrive file list ─────────────────────────────────────────────────────────

interface RawDriveItem {
  id?: string;
  name?: string;
  size?: number;
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  lastModifiedDateTime?: string;
  webUrl?: string;
}

export function sanitizeFileList(items: unknown[]): unknown[] {
  return (items as RawDriveItem[]).map(item => ({
    id: item.id,
    name: item.name,
    type: item.folder ? 'folder' : 'file',
    mimeType: item.file?.mimeType ?? null,
    size: item.size ?? null,
    childCount: item.folder?.childCount ?? null,
    modified: item.lastModifiedDateTime,
    url: item.webUrl,
  }));
}
