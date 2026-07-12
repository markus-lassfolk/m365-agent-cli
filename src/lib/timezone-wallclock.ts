/**
 * Converts a `Date` (an absolute instant) into the wall-clock date/time string for a given IANA
 * time zone, e.g. for building Microsoft Graph `dateTimeTimeZone` request fields (`{dateTime,
 * timeZone}`) where `dateTime` must be local wall-clock time in `timeZone`, not UTC/ISO.
 */

/** Throws a clean `Error` (not a raw `RangeError`) for an IANA time zone name Node/ICU doesn't recognize. */
export function assertValidTimeZone(timeZone: string): string {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone });
  } catch {
    throw new Error(`Invalid IANA time zone name: "${timeZone}" (e.g. "America/New_York", "Europe/London")`);
  }
  return timeZone;
}

/** Formats `date` as `YYYY-MM-DDTHH:mm:ss` wall-clock time in `timeZone` (no offset/Z suffix — pair with a Graph `timeZone` field). */
export function formatDateInTimeZone(date: Date, timeZone: string): string {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  }).formatToParts(date);
  const get = (type: string): string => parts.find((p) => p.type === type)?.value ?? '00';
  // Some locales render midnight as hour "24" under hour12:false; Graph expects 00-23.
  const hour = get('hour') === '24' ? '00' : get('hour');
  return `${get('year')}-${get('month')}-${get('day')}T${hour}:${get('minute')}:${get('second')}`;
}

/** Wall-clock hour (0-23) of `date` in `timeZone` — for working-hours filtering against a requested zone rather than the host's local zone. */
export function hourInTimeZone(date: Date, timeZone: string): number {
  const hour = new Intl.DateTimeFormat('en-US', { timeZone, hour: '2-digit', hour12: false })
    .formatToParts(date)
    .find((p) => p.type === 'hour')?.value;
  const n = Number(hour);
  return n === 24 ? 0 : n;
}
