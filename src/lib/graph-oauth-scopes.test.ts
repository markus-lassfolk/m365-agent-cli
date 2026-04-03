import { describe, expect, test } from 'bun:test';
import { GRAPH_DEVICE_CODE_LOGIN_SCOPES, GRAPH_REFRESH_SCOPE_CANDIDATES } from './graph-oauth-scopes.js';

describe('graph-oauth-scopes', () => {
  test('login scope string includes Shared mail/calendar and Places/People/User.Read.All', () => {
    for (const s of [
      'Mail.Send',
      'Calendars.Read.Shared',
      'Calendars.ReadWrite.Shared',
      'Mail.Read.Shared',
      'Mail.ReadWrite.Shared',
      'Place.Read.All',
      'People.Read',
      'User.Read.All',
      'MailboxSettings.ReadWrite',
      'Contacts.ReadWrite',
      'OnlineMeetings.ReadWrite',
      'Notes.ReadWrite.All'
    ]) {
      expect(GRAPH_DEVICE_CODE_LOGIN_SCOPES).toContain(s);
    }
  });

  test('refresh candidates include full Graph resource URLs and a fallback without User.Read.All', () => {
    const joined = GRAPH_REFRESH_SCOPE_CANDIDATES.join(' ');
    expect(joined).toContain('https://graph.microsoft.com/Mail.Send');
    expect(joined).toContain('https://graph.microsoft.com/User.Read.All');
    expect(joined).toContain('https://graph.microsoft.com/Place.Read.All');
    expect(joined).toContain('https://graph.microsoft.com/Mail.Read.Shared');
    expect(joined).toContain('https://graph.microsoft.com/Notes.ReadWrite.All');
    const withoutUserReadAll = GRAPH_REFRESH_SCOPE_CANDIDATES.find(
      (c) => c.includes('Mail.Read.Shared') && !c.includes('User.Read.All')
    );
    expect(withoutUserReadAll).toBeDefined();
  });
});
