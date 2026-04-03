import { describe, expect, test } from 'bun:test';
import {
  GRAPH_CRITICAL_DELEGATED_SCOPES,
  GRAPH_DEVICE_CODE_LOGIN_SCOPES,
  GRAPH_REFRESH_SCOPE_CANDIDATES
} from './graph-oauth-scopes.js';

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
      'Contacts.Read.Shared',
      'Contacts.ReadWrite.Shared',
      'OnlineMeetings.ReadWrite',
      'Notes.ReadWrite.All',
      'Team.ReadBasic.All',
      'Channel.ReadBasic.All',
      'ChannelMessage.Read.All',
      'ChannelMessage.Send',
      'Presence.Read.All',
      'Presence.ReadWrite',
      'Bookings.ReadWrite.All',
      'Chat.ReadWrite'
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
    expect(joined).toContain('https://graph.microsoft.com/Team.ReadBasic.All');
    expect(joined).toContain('https://graph.microsoft.com/ChannelMessage.Read.All');
    expect(joined).toContain('https://graph.microsoft.com/Chat.ReadWrite');
    expect(joined).toContain('https://graph.microsoft.com/ChannelMessage.Send');
    expect(GRAPH_CRITICAL_DELEGATED_SCOPES).toContain('Mail.Send');
    expect(GRAPH_CRITICAL_DELEGATED_SCOPES.length).toBeGreaterThanOrEqual(4);
    const withoutUserReadAll = GRAPH_REFRESH_SCOPE_CANDIDATES.find(
      (c) => c.includes('Mail.Read.Shared') && !c.includes('User.Read.All')
    );
    expect(withoutUserReadAll).toBeDefined();
  });
});
