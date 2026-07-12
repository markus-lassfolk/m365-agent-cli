import { mkdir, readFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import { atomicWriteUtf8File } from './atomic-write.js';
import type { DriveLocation } from './drive-location.js';

export const DELTA_STATE_VERSION = 1 as const;

export type DeltaStateKind =
  | 'mailMessages'
  | 'calendarEvents'
  | 'contacts'
  | 'todoTasks'
  | 'todoLists'
  | 'plannerAll'
  | 'driveDelta'
  | 'sharePointListItems'
  | 'meetingRecordings'
  | 'meetingTranscripts';

const VALID_KINDS = new Set<DeltaStateKind>([
  'mailMessages',
  'calendarEvents',
  'contacts',
  'todoTasks',
  'todoLists',
  'plannerAll',
  'driveDelta',
  'sharePointListItems',
  'meetingRecordings',
  'meetingTranscripts'
]);

export interface DeltaStateFileV1 {
  version: typeof DELTA_STATE_VERSION;
  kind: DeltaStateKind;
  /** ISO 8601 when the file was last written */
  updatedAt: string;
  /** After a full sync pass completes, Graph returns this for the next incremental run */
  deltaLink?: string;
  /** Mid-pagination — follow this before you have a deltaLink */
  pendingNextLink?: string;
  folderId?: string;
  calendarId?: string;
  /** To Do list id (resolved) for `todo delta`; not used for `todo lists-delta` (kind `todoLists`) */
  listId?: string;
  user?: string;
  /** `files delta` — drive root scope (mutually exclusive site fields per DriveLocation) */
  driveLocKind?: 'me' | 'user' | 'drive' | 'site';
  driveLocUser?: string;
  driveLocDriveId?: string;
  driveLocSiteId?: string;
  driveLocLibraryDriveId?: string;
  /** Folder item id when delta scope is under a folder, not drive root */
  driveFolderItemId?: string;
  /** `sharepoint items-delta` */
  sharePointSiteId?: string;
  sharePointListId?: string;
  /** `meeting recordings-all --delta` / `transcripts-all --delta` — must match before reusing cursor */
  meetingOrganizerUserId?: string;
  meetingRollupStart?: string;
  meetingRollupEnd?: string;
  meetingRollupTop?: number;
}

export function parseDeltaStateJson(raw: string): DeltaStateFileV1 | null {
  try {
    const o = JSON.parse(raw) as Partial<DeltaStateFileV1>;
    if (o.version !== 1 || !o.kind || !VALID_KINDS.has(o.kind)) {
      return null;
    }
    return o as DeltaStateFileV1;
  } catch {
    return null;
  }
}

export async function readDeltaStateFile(path: string): Promise<DeltaStateFileV1 | null> {
  try {
    const raw = await readFile(path, 'utf8');
    return parseDeltaStateJson(raw);
  } catch {
    return null;
  }
}

export async function writeDeltaStateFile(path: string, state: DeltaStateFileV1): Promise<void> {
  await mkdir(dirname(path), { recursive: true });
  // codeql[js/http-to-file-access]: Persists Graph delta cursors (next/delta links) as JSON for incremental sync; content is not executed as code.
  // Atomic temp+rename so an interrupted write can't truncate the file and lose the saved
  // delta cursor (which would force a full resync on the next run).
  await atomicWriteUtf8File(path, `${JSON.stringify(state, null, 2)}\n`, 0o644);
}

export interface DeltaScopeFields {
  folderId?: string;
  calendarId?: string;
  listId?: string;
  user?: string;
  driveLocKind?: 'me' | 'user' | 'drive' | 'site';
  driveLocUser?: string;
  driveLocDriveId?: string;
  driveLocSiteId?: string;
  driveLocLibraryDriveId?: string;
  driveFolderItemId?: string;
  sharePointSiteId?: string;
  sharePointListId?: string;
  meetingOrganizerUserId?: string;
  meetingRollupStart?: string;
  meetingRollupEnd?: string;
  meetingRollupTop?: number;
}

/** Merge a delta page response into persisted state. */
export function applyDeltaPageToState(
  prev: DeltaStateFileV1 | null,
  kind: DeltaStateKind,
  page: { '@odata.nextLink'?: string; '@odata.deltaLink'?: string },
  scope: DeltaScopeFields
): DeltaStateFileV1 {
  const nextLink = page['@odata.nextLink'];
  const deltaLink = page['@odata.deltaLink'];
  const scopeFields = {
    ...(scope.folderId !== undefined ? { folderId: scope.folderId } : {}),
    ...(scope.calendarId !== undefined ? { calendarId: scope.calendarId } : {}),
    ...(scope.listId !== undefined ? { listId: scope.listId } : {}),
    ...(scope.user !== undefined ? { user: scope.user } : {}),
    ...(scope.driveLocKind !== undefined ? { driveLocKind: scope.driveLocKind } : {}),
    ...(scope.driveLocUser !== undefined ? { driveLocUser: scope.driveLocUser } : {}),
    ...(scope.driveLocDriveId !== undefined ? { driveLocDriveId: scope.driveLocDriveId } : {}),
    ...(scope.driveLocSiteId !== undefined ? { driveLocSiteId: scope.driveLocSiteId } : {}),
    ...(scope.driveLocLibraryDriveId !== undefined ? { driveLocLibraryDriveId: scope.driveLocLibraryDriveId } : {}),
    ...(scope.driveFolderItemId !== undefined ? { driveFolderItemId: scope.driveFolderItemId } : {}),
    ...(scope.sharePointSiteId !== undefined ? { sharePointSiteId: scope.sharePointSiteId } : {}),
    ...(scope.sharePointListId !== undefined ? { sharePointListId: scope.sharePointListId } : {}),
    ...(scope.meetingOrganizerUserId !== undefined ? { meetingOrganizerUserId: scope.meetingOrganizerUserId } : {}),
    ...(scope.meetingRollupStart !== undefined ? { meetingRollupStart: scope.meetingRollupStart } : {}),
    ...(scope.meetingRollupEnd !== undefined ? { meetingRollupEnd: scope.meetingRollupEnd } : {}),
    ...(scope.meetingRollupTop !== undefined ? { meetingRollupTop: scope.meetingRollupTop } : {})
  };
  if (deltaLink) {
    return {
      version: 1,
      kind,
      updatedAt: new Date().toISOString(),
      ...scopeFields,
      deltaLink,
      pendingNextLink: undefined
    };
  }
  if (nextLink) {
    return {
      version: 1,
      kind,
      updatedAt: new Date().toISOString(),
      ...scopeFields,
      pendingNextLink: nextLink,
      deltaLink: prev?.deltaLink
    };
  }
  return {
    version: 1,
    kind,
    updatedAt: new Date().toISOString(),
    ...scopeFields,
    pendingNextLink: prev?.pendingNextLink,
    deltaLink: prev?.deltaLink
  };
}

/**
 * URL to pass as `nextLink` for the next request: explicit `--next`, then pending page, then saved delta.
 */
export function resolveDeltaContinuationUrl(opts: {
  explicitNext?: string;
  state?: DeltaStateFileV1 | null;
}): string | undefined {
  if (opts.explicitNext?.trim()) {
    return opts.explicitNext.trim();
  }
  const s = opts.state;
  if (!s) return undefined;
  if (s.pendingNextLink?.trim()) {
    return s.pendingNextLink.trim();
  }
  if (s.deltaLink?.trim()) {
    return s.deltaLink.trim();
  }
  return undefined;
}

export function assertDeltaScopeMatchesState(state: DeltaStateFileV1, scope: DeltaScopeFields): void {
  const norm = (v: string | undefined) => (v === undefined || v === '' ? undefined : v.trim().toLowerCase());
  if (state.kind === 'driveDelta') {
    if (norm(state.driveLocKind as string | undefined) !== norm(scope.driveLocKind as string | undefined)) {
      throw new Error(
        `State file drive scope (kind ${state.driveLocKind ?? ''}) does not match current drive flags (kind ${scope.driveLocKind ?? ''})`
      );
    }
    if (norm(state.driveLocUser) !== norm(scope.driveLocUser)) {
      throw new Error(
        `State file driveLocUser "${state.driveLocUser ?? ''}" does not match --user "${scope.driveLocUser ?? '(none)'}"`
      );
    }
    if (norm(state.driveLocDriveId) !== norm(scope.driveLocDriveId)) {
      throw new Error(
        `State file driveLocDriveId "${state.driveLocDriveId ?? ''}" does not match --drive-id "${scope.driveLocDriveId ?? '(none)'}"`
      );
    }
    if (norm(state.driveLocSiteId) !== norm(scope.driveLocSiteId)) {
      throw new Error(
        `State file driveLocSiteId "${state.driveLocSiteId ?? ''}" does not match --site-id "${scope.driveLocSiteId ?? '(none)'}"`
      );
    }
    if (norm(state.driveLocLibraryDriveId) !== norm(scope.driveLocLibraryDriveId)) {
      throw new Error(
        `State file driveLocLibraryDriveId "${state.driveLocLibraryDriveId ?? ''}" does not match --library-drive-id "${scope.driveLocLibraryDriveId ?? '(none)'}"`
      );
    }
    if (norm(state.driveFolderItemId) !== norm(scope.driveFolderItemId)) {
      throw new Error(
        `State file driveFolderItemId "${state.driveFolderItemId ?? ''}" does not match --folder "${scope.driveFolderItemId ?? '(none)'}"`
      );
    }
    return;
  }

  if (state.kind === 'sharePointListItems') {
    if (norm(state.sharePointSiteId) !== norm(scope.sharePointSiteId)) {
      throw new Error(
        `State file sharePointSiteId "${state.sharePointSiteId ?? ''}" does not match --site-id "${scope.sharePointSiteId ?? '(none)'}"`
      );
    }
    if (norm(state.sharePointListId) !== norm(scope.sharePointListId)) {
      throw new Error(
        `State file sharePointListId "${state.sharePointListId ?? ''}" does not match --list-id "${scope.sharePointListId ?? '(none)'}"`
      );
    }
    return;
  }

  if (state.kind === 'todoLists') {
    if (norm(state.user) !== norm(scope.user)) {
      throw new Error(`State file user "${state.user ?? ''}" does not match --user "${scope.user ?? '(none)'}"`);
    }
    return;
  }

  if (state.kind === 'meetingRecordings' || state.kind === 'meetingTranscripts') {
    const trimEq = (a: string | undefined, b: string | undefined) => (a ?? '').trim() === (b ?? '').trim();
    // Organizer id can be a UPN — compare case-insensitively (like every other identity field)
    // so `User@contoso.com` vs `user@contoso.com` doesn't reject a valid cursor. Dates stay exact.
    if (norm(state.meetingOrganizerUserId) !== norm(scope.meetingOrganizerUserId)) {
      throw new Error(
        `State file meeting organizer "${state.meetingOrganizerUserId ?? ''}" does not match current "${scope.meetingOrganizerUserId ?? '(none)'}"`
      );
    }
    if (!trimEq(state.meetingRollupStart, scope.meetingRollupStart)) {
      throw new Error(
        `State file --start "${state.meetingRollupStart ?? ''}" does not match current "${scope.meetingRollupStart ?? '(none)'}"`
      );
    }
    if (!trimEq(state.meetingRollupEnd, scope.meetingRollupEnd)) {
      throw new Error(
        `State file --end "${state.meetingRollupEnd ?? ''}" does not match current "${scope.meetingRollupEnd ?? '(none)'}"`
      );
    }
    if (
      state.meetingRollupTop !== undefined &&
      scope.meetingRollupTop !== undefined &&
      state.meetingRollupTop !== scope.meetingRollupTop
    ) {
      throw new Error(`State file --top ${state.meetingRollupTop} does not match current ${scope.meetingRollupTop}`);
    }
    if (norm(state.user) !== norm(scope.user)) {
      throw new Error(`State file user "${state.user ?? ''}" does not match --user "${scope.user ?? '(none)'}"`);
    }
    return;
  }

  if (norm(state.folderId) !== norm(scope.folderId)) {
    throw new Error(
      `State file folderId "${state.folderId ?? ''}" does not match --folder "${scope.folderId ?? '(none)'}"`
    );
  }
  if (norm(state.calendarId) !== norm(scope.calendarId)) {
    throw new Error(
      `State file calendarId "${state.calendarId ?? ''}" does not match --calendar "${scope.calendarId ?? '(none)'}"`
    );
  }
  if (norm(state.listId) !== norm(scope.listId)) {
    throw new Error(
      `State file listId "${state.listId ?? ''}" does not match resolved list "${scope.listId ?? '(none)'}"`
    );
  }
  if (norm(state.user) !== norm(scope.user)) {
    throw new Error(`State file user "${state.user ?? ''}" does not match --user "${scope.user ?? '(none)'}"`);
  }
}

/** Scope fields persisted for `files delta` (matches {@link DriveLocation} + optional folder item). */
export function driveDeltaScopeFromLocation(loc: DriveLocation, folderItemId?: string): DeltaScopeFields {
  const folder = folderItemId?.trim() || undefined;
  if (loc.kind === 'me') {
    return { driveLocKind: 'me', driveFolderItemId: folder };
  }
  if (loc.kind === 'user') {
    return { driveLocKind: 'user', driveLocUser: loc.user.trim(), driveFolderItemId: folder };
  }
  if (loc.kind === 'drive') {
    return { driveLocKind: 'drive', driveLocDriveId: loc.driveId.trim(), driveFolderItemId: folder };
  }
  const lib = loc.libraryDriveId?.trim();
  return {
    driveLocKind: 'site',
    driveLocSiteId: loc.siteId.trim(),
    ...(lib ? { driveLocLibraryDriveId: lib } : {}),
    driveFolderItemId: folder
  };
}
