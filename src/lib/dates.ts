const WEEKDAYS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'] as const;

export interface ParseDayOptions {
  baseDate?: Date;
  weekdayDirection?: 'next' | 'previous' | 'nearestForward';
  throwOnInvalid?: boolean;
}

export function parseTimeToDate(timeStr: string, baseDate: Date = new Date()): Date {
  const result = new Date(baseDate);

  const timeMatch = timeStr.match(/^(\d{1,2}):(\d{2})$/);
  if (timeMatch) {
    result.setHours(parseInt(timeMatch[1], 10), parseInt(timeMatch[2], 10), 0, 0);
    return result;
  }

  const hourMatch = timeStr.match(/^(\d{1,2})(am|pm)?$/i);
  if (hourMatch) {
    let hour = parseInt(hourMatch[1], 10);
    const isPM = hourMatch[2]?.toLowerCase() === 'pm';
    if (isPM && hour < 12) hour += 12;
    if (!isPM && hour === 12) hour = 0;
    result.setHours(hour, 0, 0, 0);
    return result;
  }

  return result;
}

export function toLocalISOString(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  const offsetMinutes = -date.getTimezoneOffset();
  const offsetSign = offsetMinutes >= 0 ? '+' : '-';
  const absMinutes = Math.abs(offsetMinutes);
  const offsetHH = String(Math.floor(absMinutes / 60)).padStart(2, '0');
  const offsetMM = String(absMinutes % 60).padStart(2, '0');
  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}${offsetSign}${offsetHH}:${offsetMM}`;
}

/**
 * Parse a YYYY-MM-DD string as local midnight (in the user's timezone),
 * not UTC midnight. This fixes off-by-one day shifts for UTC-X timezones.
 *
 * `new Date('YYYY-MM-DD')` parses as UTC midnight, which becomes the
 * previous calendar day when local timezone is west of UTC.
 */
function parseLocalDate(dayStr: string): Date {
  const match = dayStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return new Date(dayStr);
  const [, yearStr, monthStr, dayOfMonthStr] = match;
  return new Date(parseInt(yearStr, 10), parseInt(monthStr, 10) - 1, parseInt(dayOfMonthStr, 10), 0, 0, 0, 0);
}

export function parseDay(day: string, options: ParseDayOptions = {}): Date {
  const { baseDate = new Date(), weekdayDirection = 'next', throwOnInvalid = false } = options;

  const now = new Date(baseDate);
  const normalized = day.toLowerCase();

  if (normalized === 'today') return now;
  if (normalized === 'tomorrow') {
    now.setDate(now.getDate() + 1);
    return now;
  }
  if (normalized === 'yesterday') {
    now.setDate(now.getDate() - 1);
    return now;
  }

  const targetDay = WEEKDAYS.indexOf(normalized as (typeof WEEKDAYS)[number]);
  if (targetDay >= 0) {
    const currentDay = now.getDay();
    let diff = targetDay - currentDay;

    if (weekdayDirection === 'next') {
      if (diff <= 0) diff += 7;
    } else if (weekdayDirection === 'previous') {
      if (diff > 0) diff -= 7;
    } else {
      if (diff < 0) diff += 7;
    }

    now.setDate(now.getDate() + diff);
    return now;
  }

  // Parse YYYY-MM-DD as local midnight to avoid UTC off-by-one
  const parsed = parseLocalDate(day);
  if (Number.isNaN(parsed.getTime())) {
    if (throwOnInvalid) {
      throw new Error(`Invalid day value: ${day}`);
    }
    return now;
  }

  return parsed;
}
