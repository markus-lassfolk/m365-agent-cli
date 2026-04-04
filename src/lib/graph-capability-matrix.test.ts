import { describe, expect, test } from 'bun:test';
import {
  evaluateGraphCapabilities,
  GRAPH_CAPABILITY_MATRIX,
  graphTokenPermissionKind,
  permissionSetFromGraphPayload
} from './graph-capability-matrix.js';

describe('permissionSetFromGraphPayload', () => {
  test('merges scp and roles', () => {
    const s = permissionSetFromGraphPayload({
      scp: 'Mail.Read User.Read',
      roles: ['Sites.Read.All']
    });
    expect(s.has('Mail.Read')).toBe(true);
    expect(s.has('User.Read')).toBe(true);
    expect(s.has('Sites.Read.All')).toBe(true);
  });
});

describe('evaluateGraphCapabilities', () => {
  test('full login-style scopes unlock mail and calendar own', () => {
    const perms = permissionSetFromGraphPayload({
      scp: 'Calendars.ReadWrite Mail.ReadWrite Mail.Send User.Read offline_access'
    });
    const rows = evaluateGraphCapabilities(perms);
    const cal = rows.find((r) => r.id === 'calendar.own');
    const mail = rows.find((r) => r.id === 'mail.own');
    expect(cal?.readOk && cal?.writeOk).toBe(true);
    expect(mail?.readOk && mail?.writeOk).toBe(true);
  });

  test('Planner read needs Group.Read or Group.ReadWrite', () => {
    const onlyRead = evaluateGraphCapabilities(new Set(['Group.Read.All']));
    const planner = onlyRead.find((r) => r.id === 'planner.groups');
    expect(planner?.readOk).toBe(true);
    expect(planner?.writeOk).toBe(false);

    const rw = evaluateGraphCapabilities(new Set(['Group.ReadWrite.All']));
    const p2 = rw.find((r) => r.id === 'planner.groups');
    expect(p2?.readOk && p2?.writeOk).toBe(true);
  });

  test('mail.send row: Mail.Send grants write only (read column is dash in UI)', () => {
    const ev = evaluateGraphCapabilities(new Set(['Mail.Send']));
    const row = ev.find((r) => r.id === 'mail.send');
    expect(row?.writeOk).toBe(true);
    expect(row?.readOk).toBe(false);
  });

  test('SharePoint write needs Sites.ReadWrite or Sites.Manage', () => {
    const readOnly = evaluateGraphCapabilities(new Set(['Sites.Read.All']));
    const sp = readOnly.find((r) => r.id === 'sharepoint.sites');
    expect(sp?.readOk).toBe(true);
    expect(sp?.writeOk).toBe(false);

    const write = evaluateGraphCapabilities(new Set(['Sites.ReadWrite.All']));
    const sp2 = write.find((r) => r.id === 'sharepoint.sites');
    expect(sp2?.writeOk).toBe(true);
  });

  test('matrix has stable ids', () => {
    const ids = new Set(GRAPH_CAPABILITY_MATRIX.map((r) => r.id));
    expect(ids.size).toBe(GRAPH_CAPABILITY_MATRIX.length);
  });
});

describe('graphTokenPermissionKind', () => {
  test('classifies delegated vs application vs mixed', () => {
    expect(graphTokenPermissionKind({ scp: 'Mail.Read' })).toBe('delegated');
    expect(graphTokenPermissionKind({ roles: ['Mail.Read'] })).toBe('application');
    expect(graphTokenPermissionKind({ scp: 'a', roles: ['b'] })).toBe('mixed');
    expect(graphTokenPermissionKind({})).toBe('unknown');
  });
});
