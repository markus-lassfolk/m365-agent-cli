import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import type { GraphResponse } from '../lib/graph-client.js';
import {
  applyDeltaPageToState,
  assertDeltaScopeMatchesState,
  type DeltaScopeFields,
  type DeltaStateFileV1,
  readDeltaStateFile,
  resolveDeltaContinuationUrl,
  writeDeltaStateFile
} from '../lib/graph-delta-state-file.js';
import {
  type CallRecording,
  type CallTranscript,
  downloadMediaToFile,
  getAllRecordings,
  getAllTranscripts,
  getRecordingsDeltaPage,
  getTranscriptsDeltaPage,
  listMeetingRecordings,
  listMeetingTranscripts,
  recordingContentPath,
  transcriptContentPath,
  transcriptMetadataContentPath
} from '../lib/graph-meeting-recordings-client.js';
import { getUserProfile } from '../lib/graph-org-client.js';
import {
  createOnlineMeeting,
  createOnlineMeetingFromBody,
  deleteOnlineMeeting,
  getOnlineMeeting,
  type OnlineMeeting,
  updateOnlineMeeting
} from '../lib/online-meetings-graph-client.js';
import { readJsonFileOrExit } from '../lib/read-json-file.js';
import { checkReadOnly } from '../lib/utils.js';

function parseOptionalRecordingsTop(raw?: string): number | undefined {
  if (!raw?.trim()) return undefined;
  const n = Number.parseInt(raw.trim(), 10);
  if (!Number.isFinite(n) || n < 1) {
    return undefined;
  }
  return Math.min(999, n);
}

/** Resolve `me` / default to a real `meetingOrganizerUserId` for `getAllRecordings` / `getAllTranscripts`. */
async function resolveMeetingOrganizerUserId(
  token: string,
  organizerFlag: string | undefined,
  graphUserOption: string | undefined
): Promise<string | undefined> {
  const fromFlag = organizerFlag?.trim();
  if (fromFlag && fromFlag !== 'me') {
    return fromFlag;
  }
  const u = graphUserOption?.trim();
  if (u && u !== 'me') {
    return u;
  }
  const r = await getUserProfile(token);
  if (!r.ok || !r.data?.id) {
    return undefined;
  }
  return r.data.id;
}

export const meetingCommand = new Command('meeting').description(
  'Teams online meetings via Microsoft Graph (`OnlineMeetings.ReadWrite`). ' +
    '**Calendar invitations with Teams + attendees:** use `create-event ... --teams` (see `event.teamsMeeting` in `--json` output). ' +
    'This command is for **standalone** `POST /onlineMeetings` (join link without a calendar event, or advanced JSON).'
);

meetingCommand
  .command('create')
  .description(
    'Create an online meeting (`POST /me/onlineMeetings`). Use `--json-file` for full Graph body (participants, lobby, etc.).'
  )
  .option('--json-file <path>', 'Full JSON body (overrides --start/--end/--subject)')
  .option('--start <iso>', 'Start time (ISO 8601, e.g. 2026-04-03T14:00:00-07:00)')
  .option('--end <iso>', 'End time (ISO 8601)')
  .option('-s, --subject <text>', 'Meeting subject')
  .option('--json', 'Output full Graph JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation; same tenant Teams)')
  .action(
    async (
      opts: {
        jsonFile?: string;
        start?: string;
        end?: string;
        subject?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      let r: GraphResponse<OnlineMeeting>;
      if (opts.jsonFile?.trim()) {
        const body = await readJsonFileOrExit(opts.jsonFile, '--json-file');
        r = await createOnlineMeetingFromBody(auth.token!, body, opts.user);
      } else {
        if (!opts.start?.trim() || !opts.end?.trim()) {
          console.error('Error: provide --start and --end, or use --json-file with a full Graph body.');
          process.exit(1);
        }
        r = await createOnlineMeeting(
          auth.token!,
          {
            startDateTime: opts.start.trim(),
            endDateTime: opts.end.trim(),
            subject: opts.subject
          },
          opts.user
        );
      }

      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      const join = r.data.joinWebUrl ?? r.data.joinUrl;
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      if (r.data.subject) console.log(`Subject: ${r.data.subject}`);
      if (join) console.log(`Join: ${join}`);
      else console.log(JSON.stringify(r.data, null, 2));
      if (r.data.id) console.log(`Meeting id: ${r.data.id}`);
    }
  );

meetingCommand
  .command('get')
  .description('Get an online meeting by id (`GET /me/onlineMeetings/{id}`)')
  .argument('<meetingId>', 'Online meeting id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (meetingId: string, opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getOnlineMeeting(auth.token!, meetingId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify(r.data, null, 2));
    else {
      const join = r.data.joinWebUrl ?? r.data.joinUrl;
      if (r.data.subject) console.log(`Subject: ${r.data.subject}`);
      if (join) console.log(`Join: ${join}`);
      if (r.data.id) console.log(`Meeting id: ${r.data.id}`);
    }
  });

meetingCommand
  .command('update')
  .description('Update an online meeting (`PATCH /me/onlineMeetings/{id}`)')
  .argument('<meetingId>', 'Online meeting id')
  .requiredOption('--json-file <path>', 'JSON patch body per Graph')
  .option('--json', 'Echo updated meeting as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      meetingId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const patch = await readJsonFileOrExit(opts.jsonFile, '--json-file');
      const r = await updateOnlineMeeting(auth.token!, meetingId, patch, opts.user);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(r.data, null, 2));
      else {
        const join = r.data.joinWebUrl ?? r.data.joinUrl;
        if (join) console.log(`Join: ${join}`);
        console.log(`Updated meeting: ${r.data.id ?? meetingId}`);
      }
    }
  );

meetingCommand
  .command('delete')
  .description('Delete an online meeting (`DELETE /me/onlineMeetings/{id}`)')
  .argument('<meetingId>', 'Online meeting id')
  .option('--confirm', 'Confirm delete')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      meetingId: string,
      opts: { confirm?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.error('Refusing to delete without --confirm');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteOnlineMeeting(auth.token!, meetingId, opts.user);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Online meeting deleted.');
    }
  );

interface RecordingsBaseOpts {
  user?: string;
  json?: boolean;
  token?: string;
  identity?: string;
}

const recordingsBaseFlags = (cmd: Command) =>
  cmd
    .option('--user <email>', 'Target user (Graph delegation; defaults to /me)')
    .option('--json', 'Output as JSON')
    .option('--token <token>', 'Use a specific token')
    .option('--identity <name>', 'Graph token cache identity (default: default)');

recordingsBaseFlags(meetingCommand.command('recordings <meetingId>'))
  .description(
    'List recordings for a single meeting (`GET /me/onlineMeetings/{id}/recordings`). Requires `OnlineMeetingRecording.Read.All`. 403 typically means tenant Stream/Teams policy, not a CLI bug.'
  )
  .action(async (meetingId: string, opts: RecordingsBaseOpts) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listMeetingRecordings(auth.token, meetingId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message ?? 'recordings list failed'}`);
      process.exit(1);
    }
    const items = r.data.value ?? [];
    if (opts.json) {
      console.log(JSON.stringify({ value: items }, null, 2));
      return;
    }
    if (items.length === 0) {
      console.log('No recordings.');
      return;
    }
    for (const it of items) renderRecording(it);
  });

recordingsBaseFlags(meetingCommand.command('recording-download <meetingId> <recordingId>'))
  .description(
    'Download a meeting recording (`GET /me/onlineMeetings/{id}/recordings/{id}/content`). Streams to disk; follows a single redirect into Microsoft Stream/SharePoint.'
  )
  .option('--out <path>', 'Output file path (default ./<recordingId>.mp4)')
  .action(async (meetingId: string, recordingId: string, opts: RecordingsBaseOpts & { out?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const path = recordingContentPath(meetingId, recordingId, opts.user);
    const out = opts.out?.trim() || `./${recordingId}.mp4`;
    const r = await downloadMediaToFile(auth.token, path, out);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message ?? 'download failed'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    console.log(`✓ Downloaded ${r.data.bytes} bytes`);
    console.log(`  Saved to: ${r.data.path}`);
  });

recordingsBaseFlags(meetingCommand.command('recordings-all'))
  .description(
    'Tenant-wide / per-organizer recordings (`getAllRecordings(...)`). With `--delta`, uses `getAllRecordings(...)/delta` and supports `--state-file`. Requires `OnlineMeetingRecording.Read.All`.'
  )
  .option('--organizer <upn-or-id>', 'Meeting organizer (defaults to the signed-in user)')
  .option(
    '--start <iso>',
    'Start window (required for non-delta unless --next). With --delta, optional — omit both --start and --end for a full initial sync (organizer only).'
  )
  .option(
    '--end <iso>',
    'End window (required for non-delta unless --next). With --delta, optional — if set, --start is required.'
  )
  .option('--next <url>', 'Follow `@odata.nextLink` from a previous page')
  .option('--delta', 'Use `getAllRecordings(...)/delta` for incremental sync')
  .option('--state-file <path>', '(With --delta) read/write JSON delta cursor (kind: meetingRecordings)')
  .option('--top <n>', 'Limit per page (Graph $top) — applies to non-delta calls')
  .action(
    async (
      opts: RecordingsBaseOpts & {
        organizer?: string;
        start?: string;
        end?: string;
        next?: string;
        delta?: boolean;
        stateFile?: string;
        top?: string;
      }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const accessToken = auth.token;
      if (opts.delta) {
        const existingForCont = opts.stateFile?.trim() ? await readDeltaStateFile(opts.stateFile.trim()) : null;
        const organizerUserId = await resolveMeetingOrganizerUserId(accessToken, opts.organizer, opts.user);
        if (!organizerUserId) {
          console.error('Error: could not resolve meeting organizer (try --organizer <upn-or-id>).');
          process.exit(1);
        }
        const topParsed = parseOptionalRecordingsTop(opts.top);
        const rollupStart = (opts.start?.trim() || existingForCont?.meetingRollupStart || '').trim();
        const rollupEnd = (opts.end?.trim() || existingForCont?.meetingRollupEnd || '').trim();
        const rollupTop = topParsed ?? existingForCont?.meetingRollupTop;
        const meetingDeltaScope: DeltaScopeFields = {
          user: opts.user,
          meetingOrganizerUserId: organizerUserId,
          meetingRollupStart: rollupStart,
          meetingRollupEnd: rollupEnd,
          ...(rollupTop !== undefined ? { meetingRollupTop: rollupTop } : {})
        };
        await runMeetingDelta({
          auth: accessToken,
          kind: 'meetingRecordings',
          stateFile: opts.stateFile,
          explicitNext: opts.next,
          json: opts.json,
          fetchPage: (pageUrl) =>
            getRecordingsDeltaPage(accessToken, {
              pageUrl,
              user: opts.user,
              initial: pageUrl?.trim()
                ? undefined
                : {
                    organizerUserId,
                    startDateTime: rollupStart,
                    endDateTime: rollupEnd,
                    top: rollupTop
                  }
            }),
          renderItem: (it: CallRecording) => renderRecording(it),
          scope: meetingDeltaScope
        });
        return;
      }
      if (!opts.start?.trim() || !opts.end?.trim()) {
        if (!opts.next?.trim()) {
          console.error('Error: --start and --end are required (unless following --next or using --delta)');
          process.exit(1);
        }
      }
      const organizerUserId = await resolveMeetingOrganizerUserId(accessToken, opts.organizer, opts.user);
      if (!organizerUserId) {
        console.error('Error: could not resolve meeting organizer (try --organizer <upn-or-id>).');
        process.exit(1);
      }
      const topParsed = parseOptionalRecordingsTop(opts.top);
      const r = await getAllRecordings(accessToken, {
        organizerUserId,
        start: opts.start ?? '',
        end: opts.end ?? '',
        user: opts.user,
        pageUrl: opts.next,
        top: topParsed
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message ?? 'getAllRecordings failed'}`);
        process.exit(1);
      }
      const items = r.data.value ?? [];
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const it of items) renderRecording(it);
      if (r.data['@odata.nextLink']) console.log(`nextLink: ${r.data['@odata.nextLink']}`);
    }
  );

recordingsBaseFlags(meetingCommand.command('transcripts <meetingId>'))
  .description(
    'List transcripts for a single meeting (`GET /me/onlineMeetings/{id}/transcripts`). Requires `OnlineMeetingTranscript.Read.All`.'
  )
  .action(async (meetingId: string, opts: RecordingsBaseOpts) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listMeetingTranscripts(auth.token, meetingId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message ?? 'transcripts list failed'}`);
      process.exit(1);
    }
    const items = r.data.value ?? [];
    if (opts.json) {
      console.log(JSON.stringify({ value: items }, null, 2));
      return;
    }
    if (items.length === 0) {
      console.log('No transcripts.');
      return;
    }
    for (const it of items) renderTranscript(it);
  });

recordingsBaseFlags(meetingCommand.command('transcript-download <meetingId> <transcriptId>'))
  .description(
    'Download a transcript (`GET /me/onlineMeetings/{id}/transcripts/{id}/content` returns VTT). Pass `--metadata` to also fetch `metadataContent` (utterance timing).'
  )
  .option('--out <path>', 'Output VTT path (default ./<transcriptId>.vtt)')
  .option('--metadata', 'Also download `metadataContent` to <out>.metadata.json')
  .action(
    async (
      meetingId: string,
      transcriptId: string,
      opts: RecordingsBaseOpts & { out?: string; metadata?: boolean }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const out = opts.out?.trim() || `./${transcriptId}.vtt`;
      const r = await downloadMediaToFile(auth.token, transcriptContentPath(meetingId, transcriptId, opts.user), out);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message ?? 'transcript download failed'}`);
        process.exit(1);
      }
      let metaResult: { path: string; bytes: number } | undefined;
      if (opts.metadata) {
        const metaOut = `${out}.metadata.json`;
        const m = await downloadMediaToFile(
          auth.token,
          transcriptMetadataContentPath(meetingId, transcriptId, opts.user),
          metaOut
        );
        if (!m.ok || !m.data) {
          console.error(`Warning: metadata download failed: ${m.error?.message ?? '?'}`);
        } else {
          metaResult = m.data;
        }
      }
      if (opts.json) {
        console.log(JSON.stringify({ content: r.data, metadata: metaResult ?? null }, null, 2));
        return;
      }
      console.log(`✓ Downloaded transcript ${r.data.bytes} bytes`);
      console.log(`  Saved to: ${r.data.path}`);
      if (metaResult) console.log(`  Metadata: ${metaResult.path}`);
    }
  );

recordingsBaseFlags(meetingCommand.command('transcripts-all'))
  .description(
    'Tenant-wide / per-organizer transcripts (`getAllTranscripts(...)`). With `--delta`, uses `getAllTranscripts(...)/delta` + `--state-file`. Requires `OnlineMeetingTranscript.Read.All`.'
  )
  .option('--organizer <upn-or-id>', 'Meeting organizer (defaults to the signed-in user)')
  .option(
    '--start <iso>',
    'Start window (required for non-delta unless --next). With --delta, optional — omit both --start and --end for a full initial sync (organizer only).'
  )
  .option(
    '--end <iso>',
    'End window (required for non-delta unless --next). With --delta, optional — if set, --start is required.'
  )
  .option('--next <url>', 'Follow `@odata.nextLink` from a previous page')
  .option('--delta', 'Use `getAllTranscripts(...)/delta` for incremental sync')
  .option('--state-file <path>', '(With --delta) read/write JSON delta cursor (kind: meetingTranscripts)')
  .option('--top <n>', 'Limit per page (Graph $top) — applies to non-delta calls')
  .action(
    async (
      opts: RecordingsBaseOpts & {
        organizer?: string;
        start?: string;
        end?: string;
        next?: string;
        delta?: boolean;
        stateFile?: string;
        top?: string;
      }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const accessToken = auth.token;
      if (opts.delta) {
        const existingForCont = opts.stateFile?.trim() ? await readDeltaStateFile(opts.stateFile.trim()) : null;
        const organizerUserId = await resolveMeetingOrganizerUserId(accessToken, opts.organizer, opts.user);
        if (!organizerUserId) {
          console.error('Error: could not resolve meeting organizer (try --organizer <upn-or-id>).');
          process.exit(1);
        }
        const topParsed = parseOptionalRecordingsTop(opts.top);
        const rollupStart = (opts.start?.trim() || existingForCont?.meetingRollupStart || '').trim();
        const rollupEnd = (opts.end?.trim() || existingForCont?.meetingRollupEnd || '').trim();
        const rollupTop = topParsed ?? existingForCont?.meetingRollupTop;
        const meetingDeltaScope: DeltaScopeFields = {
          user: opts.user,
          meetingOrganizerUserId: organizerUserId,
          meetingRollupStart: rollupStart,
          meetingRollupEnd: rollupEnd,
          ...(rollupTop !== undefined ? { meetingRollupTop: rollupTop } : {})
        };
        await runMeetingDelta({
          auth: accessToken,
          kind: 'meetingTranscripts',
          stateFile: opts.stateFile,
          explicitNext: opts.next,
          json: opts.json,
          fetchPage: (pageUrl) =>
            getTranscriptsDeltaPage(accessToken, {
              pageUrl,
              user: opts.user,
              initial: pageUrl?.trim()
                ? undefined
                : {
                    organizerUserId,
                    startDateTime: rollupStart,
                    endDateTime: rollupEnd,
                    top: rollupTop
                  }
            }),
          renderItem: (it: CallTranscript) => renderTranscript(it),
          scope: meetingDeltaScope
        });
        return;
      }
      if (!opts.start?.trim() || !opts.end?.trim()) {
        if (!opts.next?.trim()) {
          console.error('Error: --start and --end are required (unless following --next or using --delta)');
          process.exit(1);
        }
      }
      const organizerUserId = await resolveMeetingOrganizerUserId(accessToken, opts.organizer, opts.user);
      if (!organizerUserId) {
        console.error('Error: could not resolve meeting organizer (try --organizer <upn-or-id>).');
        process.exit(1);
      }
      const topParsed = parseOptionalRecordingsTop(opts.top);
      const r = await getAllTranscripts(accessToken, {
        organizerUserId,
        start: opts.start ?? '',
        end: opts.end ?? '',
        user: opts.user,
        pageUrl: opts.next,
        top: topParsed
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message ?? 'getAllTranscripts failed'}`);
        process.exit(1);
      }
      const items = r.data.value ?? [];
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const it of items) renderTranscript(it);
      if (r.data['@odata.nextLink']) console.log(`nextLink: ${r.data['@odata.nextLink']}`);
    }
  );

function renderRecording(it: CallRecording): void {
  console.log(`recording: ${it.id}`);
  if (it.meetingId) console.log(`  meetingId: ${it.meetingId}`);
  if (it.createdDateTime) console.log(`  created: ${it.createdDateTime}`);
  if (it.endDateTime) console.log(`  end: ${it.endDateTime}`);
  if (it.recordingContentUrl) console.log(`  contentUrl: ${it.recordingContentUrl}`);
}

function renderTranscript(it: CallTranscript): void {
  console.log(`transcript: ${it.id}`);
  if (it.meetingId) console.log(`  meetingId: ${it.meetingId}`);
  if (it.createdDateTime) console.log(`  created: ${it.createdDateTime}`);
  if (it.endDateTime) console.log(`  end: ${it.endDateTime}`);
  if (it.transcriptContentUrl) console.log(`  contentUrl: ${it.transcriptContentUrl}`);
}

interface MeetingDeltaArgs<T> {
  auth: string;
  kind: 'meetingRecordings' | 'meetingTranscripts';
  stateFile?: string;
  explicitNext?: string;
  json?: boolean;
  fetchPage: (pageUrl?: string) => Promise<
    GraphResponse<{
      value?: T[];
      '@odata.nextLink'?: string;
      '@odata.deltaLink'?: string;
    }>
  >;
  renderItem: (it: T) => void;
  scope: DeltaScopeFields;
}

async function runMeetingDelta<T>(args: MeetingDeltaArgs<T>): Promise<void> {
  const existing: DeltaStateFileV1 | null = args.stateFile ? await readDeltaStateFile(args.stateFile) : null;
  if (existing && existing.kind !== args.kind) {
    console.error(`Error: state file kind '${existing.kind}' does not match expected '${args.kind}'`);
    process.exit(1);
  }
  if (existing) {
    try {
      assertDeltaScopeMatchesState(existing, args.scope);
    } catch (e) {
      console.error(`Error: ${e instanceof Error ? e.message : String(e)}`);
      process.exit(1);
    }
  }
  const continueUrl = resolveDeltaContinuationUrl({ explicitNext: args.explicitNext, state: existing });
  const r = await args.fetchPage(continueUrl);
  if (!r.ok || !r.data) {
    console.error(`Error: ${r.error?.message ?? 'delta page failed'}`);
    process.exit(1);
  }
  if (args.stateFile) {
    const merged = applyDeltaPageToState(existing, args.kind, r.data, args.scope);
    await writeDeltaStateFile(args.stateFile, merged);
  }
  if (args.json) {
    console.log(JSON.stringify(r.data, null, 2));
    return;
  }
  const items = r.data.value ?? [];
  for (const it of items) args.renderItem(it);
  console.log(`Changes: ${items.length} item(s)`);
  if (r.data['@odata.nextLink']) console.log(`nextLink: ${r.data['@odata.nextLink']}`);
  if (r.data['@odata.deltaLink']) console.log(`deltaLink: ${r.data['@odata.deltaLink']}`);
  if (args.stateFile) console.log(`state-file: ${args.stateFile} (updated)`);
}
