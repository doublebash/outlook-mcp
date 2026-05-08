// Calendar-specific helpers, currently the recurrence-pattern translator.
// Microsoft Graph requires a verbose nested {pattern, range} object for
// recurring events; this helper accepts a flatter, more LLM-friendly input.

export interface RecurrenceInput {
  pattern: 'daily' | 'weekly' | 'monthly' | 'yearly';
  interval?: number;                                       // default 1
  days_of_week?: string[];                                 // for weekly: ["monday", ...]
  day_of_month?: number;                                   // for monthly/yearly: 1-31
  month?: number;                                          // for yearly: 1-12
  end_date?: string;                                       // ISO date "YYYY-MM-DD"
  occurrences?: number;                                    // alternative to end_date
}

interface GraphRecurrence {
  pattern: Record<string, unknown>;
  range: Record<string, unknown>;
}

const VALID_DAYS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];

export function buildRecurrence(
  input: RecurrenceInput,
  startDateTime: string,
  timezone: string
): GraphRecurrence {
  if (!input.pattern) {
    throw new Error('recurrence.pattern is required (daily, weekly, monthly, or yearly)');
  }
  const interval = input.interval ?? 1;
  if (interval < 1) throw new Error('recurrence.interval must be >= 1');

  const pattern: Record<string, unknown> = { interval };

  switch (input.pattern) {
    case 'daily':
      pattern.type = 'daily';
      break;

    case 'weekly': {
      pattern.type = 'weekly';
      if (!input.days_of_week || input.days_of_week.length === 0) {
        throw new Error('recurrence.days_of_week required for weekly pattern (e.g. ["monday","wednesday"])');
      }
      const days = input.days_of_week.map(d => d.toLowerCase());
      for (const d of days) {
        if (!VALID_DAYS.includes(d)) throw new Error(`Invalid day "${d}". Use: ${VALID_DAYS.join(', ')}`);
      }
      pattern.daysOfWeek = days;
      pattern.firstDayOfWeek = 'monday';
      break;
    }

    case 'monthly':
      pattern.type = 'absoluteMonthly';
      if (!input.day_of_month || input.day_of_month < 1 || input.day_of_month > 31) {
        throw new Error('recurrence.day_of_month required for monthly pattern (1-31)');
      }
      pattern.dayOfMonth = input.day_of_month;
      break;

    case 'yearly':
      pattern.type = 'absoluteYearly';
      if (!input.day_of_month || !input.month) {
        throw new Error('recurrence.day_of_month and recurrence.month both required for yearly pattern');
      }
      if (input.month < 1 || input.month > 12) throw new Error('recurrence.month must be 1-12');
      pattern.dayOfMonth = input.day_of_month;
      pattern.month = input.month;
      break;

    default:
      throw new Error(`Unknown recurrence pattern: ${input.pattern}`);
  }

  // Range — use endDate, numbered, or noEnd. Default to noEnd if neither provided.
  const startDate = startDateTime.split('T')[0]; // YYYY-MM-DD
  const range: Record<string, unknown> = {
    startDate,
    recurrenceTimeZone: timezone,
  };

  if (input.end_date && input.occurrences) {
    throw new Error('Provide either recurrence.end_date OR recurrence.occurrences, not both');
  }
  if (input.end_date) {
    range.type = 'endDate';
    range.endDate = input.end_date;
  } else if (input.occurrences) {
    if (input.occurrences < 1) throw new Error('recurrence.occurrences must be >= 1');
    range.type = 'numbered';
    range.numberOfOccurrences = input.occurrences;
  } else {
    range.type = 'noEnd';
  }

  return { pattern, range };
}
