/**
 * Graph often returns `dateTime` without a zone suffix and `timeZone` separately (commonly `UTC`).
 * Without normalization, `new Date(dateTime)` treats unzoned strings as local time and shifts UTC values.
 */
export function normalizeGraphDateTimeForParsing(dateTime: string | undefined, timeZone: string | undefined): string {
  if (!dateTime) return '';
  const dt = dateTime.trim();
  const hasExplicitZone = /(?:Z|[+-]\d{2}:?\d{2})$/i.test(dt);
  if (hasExplicitZone) return dt;
  const tz = (timeZone ?? '').trim();
  if (tz.toUpperCase() === 'UTC') {
    return `${dt}Z`;
  }
  return dt;
}
