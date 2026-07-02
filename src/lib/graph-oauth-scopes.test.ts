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
      'Tasks.Read.Shared',
      'Tasks.ReadWrite.Shared',
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
      'Chat.ReadWrite',
      'ExternalItem.Read.All',
      'Reports.Read.All',
      'CopilotPackages.Read.All',
      'CopilotPackages.ReadWrite.All',
      'OnlineMeetingAiInsight.Read.All',
      'OnlineMeetingTranscript.Read.All',
      'People.Read.All',
      'AiEnterpriseInteraction.Read',
      'LearningAssignedCourse.Read',
      'EngagementRole.Read.All',
      'EngagementRole.ReadWrite.All'
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
    expect(joined).toContain('https://graph.microsoft.com/Tasks.Read.Shared');
    expect(joined).toContain('https://graph.microsoft.com/Tasks.ReadWrite.Shared');
    expect(joined).toContain('https://graph.microsoft.com/Team.ReadBasic.All');
    expect(joined).toContain('https://graph.microsoft.com/ChannelMessage.Read.All');
    expect(joined).toContain('https://graph.microsoft.com/Chat.ReadWrite');
    expect(joined).toContain('https://graph.microsoft.com/ExternalItem.Read.All');
    expect(joined).toContain('https://graph.microsoft.com/Reports.Read.All');
    expect(joined).toContain('https://graph.microsoft.com/CopilotPackages.Read.All');
    expect(joined).toContain('https://graph.microsoft.com/OnlineMeetingAiInsight.Read.All');
    expect(joined).toContain('https://graph.microsoft.com/ChannelMessage.Send');
    expect(GRAPH_CRITICAL_DELEGATED_SCOPES).toContain('Mail.Send');
    expect(GRAPH_CRITICAL_DELEGATED_SCOPES).toContain('Tasks.ReadWrite.Shared');
    expect(GRAPH_CRITICAL_DELEGATED_SCOPES.length).toBeGreaterThanOrEqual(4);
    const withoutUserReadAll = GRAPH_REFRESH_SCOPE_CANDIDATES.find(
      (c) => c.includes('Mail.Read.Shared') && !c.includes('User.Read.All')
    );
    expect(withoutUserReadAll).toBeDefined();
  });

  test('M-3: Viva/Engage scopes appear in both login source-of-truth and refresh candidates', () => {
    // Regression for the audit M-3 finding: EngagementRole.Read.All appeared in refresh
    // scope candidates but NOT in GRAPH_DEVICE_CODE_LOGIN_SCOPES, so a brand-new
    // device-code login would never request it from the user.
    const loginScopes = new Set(GRAPH_DEVICE_CODE_LOGIN_SCOPES.split(/\s+/).filter(Boolean));
    const refreshScopesJoined = GRAPH_REFRESH_SCOPE_CANDIDATES.join(' ');

    // The audit-listed scopes must appear in BOTH places.
    for (const s of ['EngagementRole.Read.All', 'EngagementRole.ReadWrite.All', 'LearningAssignedCourse.Read']) {
      expect(loginScopes.has(s)).toBe(true);
      expect(refreshScopesJoined).toContain(`https://graph.microsoft.com/${s}`);
    }

    // Additionally, every "primary" refresh scope (the broad resource URLs, not narrow
    // fallbacks like `Files.ReadWrite` or `.default`) should be requested at login so
    // the device-code grant covers them. Pull those out and verify.
    const allShortScopes = new Set<string>();
    for (const candidate of GRAPH_REFRESH_SCOPE_CANDIDATES) {
      for (const seg of candidate.split(/\s+/)) {
        if (!seg) continue;
        const m = seg.match(/^https:\/\/graph\.microsoft\.com\/([A-Za-z0-9.]+)$/);
        if (m) allShortScopes.add(m[1]);
      }
    }
    // Narrow fallback / placeholder scopes intentionally not in the broad login grant.
    // The point of these candidates is to keep refresh working on tenants that did not
    // grant the broader ones. The login scope string remains the source of truth.
    const narrowFallbackAllowList = new Set([
      '.default',
      'Files.ReadWrite', // narrower than Files.ReadWrite.All
      'Files.Read' // narrower still
    ]);
    const missingPrimary = [...allShortScopes].filter((s) => !loginScopes.has(s) && !narrowFallbackAllowList.has(s));
    expect(missingPrimary).toEqual([]);
  });
});
