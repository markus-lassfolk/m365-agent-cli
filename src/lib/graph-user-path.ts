/**
 * Build `/me/...` or `/users/{upnOrId}/...` paths for Microsoft Graph delegation.
 *
 * @param user - User UPN, SMTP address, or object ID. Omit or empty for `/me`.
 * @param suffix - Path after `me` or `users/{id}` (no leading slash), e.g. `mailboxSettings`, `calendar/getSchedule`.
 */
export function graphUserPath(user: string | undefined, suffix: string): string {
  const s = suffix.replace(/^\//, '');
  if (!user?.trim()) {
    return `/me/${s}`;
  }
  return `/users/${encodeURIComponent(user.trim())}/${s}`;
}
