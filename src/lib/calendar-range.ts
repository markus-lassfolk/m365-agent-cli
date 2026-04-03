/** Helpers for calendar date windows (business days vs calendar days). */

export function isWeekend(d: Date): boolean {
  const day = d.getDay();
  return day === 0 || day === 6;
}

const WEEK_KEYWORDS = new Set(['week', 'thisweek', 'lastweek', 'nextweek']);

export function isWeekRangeKeyword(startDay: string): boolean {
  return WEEK_KEYWORDS.has(startDay.toLowerCase().trim());
}

/** Forward N calendar days inclusive from anchor (first day = anchor at local midnight). */
export function calendarDaysForward(anchor: Date, n: number): { start: Date; endExclusive: Date } {
  if (!Number.isFinite(n) || n < 1) {
    throw new Error('--days must be a positive integer');
  }
  const start = new Date(anchor);
  start.setHours(0, 0, 0, 0);
  const lastDay = new Date(start);
  lastDay.setDate(lastDay.getDate() + n - 1);
  const endExclusive = new Date(lastDay);
  endExclusive.setDate(endExclusive.getDate() + 1);
  endExclusive.setHours(0, 0, 0, 0);
  return { start, endExclusive };
}

/** Last N calendar days inclusive ending on anchor (anchor day included). */
export function calendarDaysBackward(anchor: Date, n: number): { start: Date; endExclusive: Date } {
  if (!Number.isFinite(n) || n < 1) {
    throw new Error('--previous-days must be a positive integer');
  }
  const endDay = new Date(anchor);
  endDay.setHours(0, 0, 0, 0);
  const start = new Date(endDay);
  start.setDate(start.getDate() - (n - 1));
  const endExclusive = new Date(endDay);
  endExclusive.setDate(endExclusive.getDate() + 1);
  endExclusive.setHours(0, 0, 0, 0);
  return { start, endExclusive };
}

/** N weekdays (Mon–Fri) inclusive forward; if anchor is Sat/Sun, first counted day is next Monday. */
export function businessDaysForward(anchor: Date, n: number): { start: Date; endExclusive: Date } {
  if (!Number.isFinite(n) || n < 1) {
    throw new Error('--business-days must be a positive integer');
  }
  const cur = new Date(anchor);
  cur.setHours(0, 0, 0, 0);
  while (isWeekend(cur)) {
    cur.setDate(cur.getDate() + 1);
  }
  const start = new Date(cur);
  let count = 0;
  while (count < n) {
    if (!isWeekend(cur)) {
      count++;
    }
    if (count === n) {
      break;
    }
    cur.setDate(cur.getDate() + 1);
  }
  const endExclusive = new Date(cur);
  endExclusive.setDate(endExclusive.getDate() + 1);
  endExclusive.setHours(0, 0, 0, 0);
  return { start, endExclusive };
}

/** N weekdays inclusive looking backward from anchor; if anchor is Sat/Sun, range ends on previous Friday. */
export function businessDaysBackward(anchor: Date, n: number): { start: Date; endExclusive: Date } {
  if (!Number.isFinite(n) || n < 1) {
    throw new Error('--previous-business-days must be a positive integer');
  }
  const cur = new Date(anchor);
  cur.setHours(0, 0, 0, 0);
  while (isWeekend(cur)) {
    cur.setDate(cur.getDate() - 1);
  }
  let count = 0;
  let start = new Date(cur);
  while (count < n) {
    if (!isWeekend(cur)) {
      count++;
      start = new Date(cur);
    }
    if (count === n) {
      break;
    }
    cur.setDate(cur.getDate() - 1);
  }
  start.setHours(0, 0, 0, 0);
  const endDay = new Date(anchor);
  endDay.setHours(0, 0, 0, 0);
  while (isWeekend(endDay)) {
    endDay.setDate(endDay.getDate() - 1);
  }
  const endExclusive = new Date(endDay);
  endExclusive.setDate(endExclusive.getDate() + 1);
  endExclusive.setHours(0, 0, 0, 0);
  return { start, endExclusive };
}

/**
 * Raise the query window start to `at` (default: current time) when it would otherwise begin earlier,
 * so calendar APIs only return events overlapping [start, end) — omitting meetings that already ended
 * while keeping ongoing and future items.
 */
export function clipCalendarRangeStartToNow(
  range: { start: string; end: string; label: string },
  at: Date = new Date()
): { start: string; end: string; label: string } {
  const now = at.getTime();
  const rangeStart = new Date(range.start).getTime();
  const rangeEnd = new Date(range.end).getTime();
  if (now >= rangeEnd) {
    throw new Error('The selected time range already ended; nothing to show.');
  }
  const apiStart = Math.max(now, rangeStart);
  if (apiStart >= rangeEnd) {
    throw new Error('Nothing remaining in this range from the current time.');
  }
  const moved = apiStart > rangeStart;
  return {
    start: new Date(apiStart).toISOString(),
    end: range.end,
    label: moved ? `${range.label} (from now)` : range.label
  };
}
