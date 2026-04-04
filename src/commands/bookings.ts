import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  cancelBookingAppointment,
  createBookingAppointment,
  createBookingCustomer,
  createBookingCustomQuestion,
  createBookingService,
  createBookingStaffMember,
  deleteBookingAppointment,
  deleteBookingCustomer,
  deleteBookingCustomQuestion,
  deleteBookingService,
  deleteBookingStaffMember,
  getBookingAppointment,
  getBookingBusiness,
  getBookingCustomer,
  getBookingService,
  getBookingStaffAvailability,
  getBookingStaffMember,
  listBookingAppointments,
  listBookingBusinesses,
  listBookingCalendarView,
  listBookingCurrencies,
  listBookingCustomers,
  listBookingCustomQuestions,
  listBookingServices,
  listBookingStaffMembers,
  updateBookingAppointment,
  updateBookingBusiness,
  updateBookingCustomer,
  updateBookingCustomQuestion,
  updateBookingService,
  updateBookingStaffMember
} from '../lib/graph-bookings-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const bookingsCommand = new Command('bookings').description(
  'Microsoft Bookings (Graph): read + write (`Bookings.ReadWrite.All`) — businesses, appointments, customers, services, staff, custom questions, calendar (see GRAPH_SCOPES.md)'
);

bookingsCommand
  .command('businesses')
  .description('List booking businesses (GET /solutions/bookingBusinesses)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listBookingBusinesses(auth.token);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const b of r.data) {
      console.log(`${b.displayName ?? '(business)'}\t${b.id}`);
    }
  });

bookingsCommand
  .command('business-get')
  .description('Get one booking business by id (GET /solutions/bookingBusinesses/{id})')
  .argument('<businessId>', 'Booking business id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (businessId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getBookingBusiness(auth.token, businessId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(
      opts.json
        ? JSON.stringify(r.data, null, 2)
        : `${r.data.displayName ?? '(business)'}\t${r.data.id}\t${r.data.businessType ?? ''}`
    );
  });

bookingsCommand
  .command('currencies')
  .description('List supported booking currency codes (GET /solutions/bookingCurrencies)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listBookingCurrencies(auth.token);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const c of r.data) {
      console.log(`${c.id ?? ''}\t${c.symbol ?? ''}`);
    }
  });

bookingsCommand
  .command('appointments')
  .description('List appointments for a booking business')
  .argument('<businessId>', 'Booking business id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (businessId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listBookingAppointments(auth.token, businessId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const a of r.data) {
      const start = a.start?.dateTime ?? '';
      const end = a.end?.dateTime ?? '';
      const who = a.customers?.[0]?.name ?? a.customers?.[0]?.emailAddress ?? '';
      console.log(`${start}\t${end}\t${who}\t${a.id ?? ''}`);
    }
  });

bookingsCommand
  .command('services')
  .description('List services for a booking business')
  .argument('<businessId>', 'Booking business id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (businessId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listBookingServices(auth.token, businessId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const s of r.data) {
      console.log(`${s.displayName ?? '(service)'}\t${s.defaultDuration ?? ''}\t${s.id ?? ''}`);
    }
  });

bookingsCommand
  .command('service-get')
  .description('Get one booking service (GET …/bookingBusinesses/{id}/services/{id})')
  .argument('<businessId>', 'Booking business id')
  .argument('<serviceId>', 'Service id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (businessId: string, serviceId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getBookingService(auth.token, businessId, serviceId);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(
        opts.json
          ? JSON.stringify(r.data, null, 2)
          : `${r.data.displayName ?? '(service)'}\t${r.data.defaultDuration ?? ''}\t${r.data.id ?? ''}`
      );
    }
  );

bookingsCommand
  .command('staff')
  .description('List staff members for a booking business')
  .argument('<businessId>', 'Booking business id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (businessId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listBookingStaffMembers(auth.token, businessId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const s of r.data) {
      console.log(`${s.displayName ?? '(staff)'}\t${s.emailAddress ?? ''}\t${s.role ?? ''}\t${s.id ?? ''}`);
    }
  });

bookingsCommand
  .command('staff-get')
  .description('Get one booking staff member (GET …/bookingBusinesses/{id}/staffMembers/{id})')
  .argument('<businessId>', 'Booking business id')
  .argument('<staffId>', 'Staff member id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (businessId: string, staffId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getBookingStaffMember(auth.token, businessId, staffId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(
      opts.json
        ? JSON.stringify(r.data, null, 2)
        : `${r.data.displayName ?? '(staff)'}\t${r.data.emailAddress ?? ''}\t${r.data.role ?? ''}\t${r.data.id ?? ''}`
    );
  });

bookingsCommand
  .command('calendar-view')
  .description('Appointments in a time window (GET …/calendarView?start=&end=; ISO 8601 times)')
  .argument('<businessId>', 'Booking business id')
  .requiredOption('--start <iso>', 'Window start e.g. 2026-04-01T00:00:00Z')
  .requiredOption('--end <iso>', 'Window end e.g. 2026-04-07T23:59:59Z')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      opts: { start: string; end: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await listBookingCalendarView(auth.token, businessId, opts.start, opts.end);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      for (const a of r.data) {
        const start = a.start?.dateTime ?? '';
        const end = a.end?.dateTime ?? '';
        const who = a.customers?.[0]?.name ?? a.customers?.[0]?.emailAddress ?? '';
        console.log(`${start}\t${end}\t${who}\t${a.id ?? ''}`);
      }
    }
  );

bookingsCommand
  .command('customers')
  .description('List customers for a booking business')
  .argument('<businessId>', 'Booking business id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (businessId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listBookingCustomers(auth.token, businessId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const c of r.data) {
      console.log(`${c.displayName ?? '(customer)'}\t${c.emailAddress ?? ''}\t${c.id ?? ''}`);
    }
  });

bookingsCommand
  .command('customer')
  .description('Get one customer by id')
  .argument('<businessId>', 'Booking business id')
  .argument('<customerId>', 'Customer id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (businessId: string, customerId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getBookingCustomer(auth.token, businessId, customerId);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(
        opts.json
          ? JSON.stringify(r.data, null, 2)
          : `${r.data.displayName ?? ''}\t${r.data.emailAddress ?? ''}\t${r.data.id ?? ''}`
      );
    }
  );

bookingsCommand
  .command('custom-questions')
  .description('List custom booking questions (GET …/customQuestions)')
  .argument('<businessId>', 'Booking business id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (businessId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listBookingCustomQuestions(auth.token, businessId);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const q of r.data) {
      const optsStr = (q.answerOptions ?? []).join(';');
      console.log(`${q.displayName ?? '(question)'}\t${q.answerInputType ?? ''}\t${optsStr}\t${q.id ?? ''}`);
    }
  });

bookingsCommand
  .command('appointment')
  .description('Get one appointment by id (richer fields than list)')
  .argument('<businessId>', 'Booking business id')
  .argument('<appointmentId>', 'Appointment id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (businessId: string, appointmentId: string, opts: { json?: boolean; token?: string; identity?: string }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await getBookingAppointment(auth.token, businessId, appointmentId);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      const a = r.data;
      const start = a.start?.dateTime ?? '';
      const end = a.end?.dateTime ?? '';
      const who = a.customerName ?? a.customers?.[0]?.name ?? a.customers?.[0]?.emailAddress ?? '';
      console.log(`${start}\t${end}\t${who}\t${a.id ?? ''}`);
    }
  );

bookingsCommand
  .command('business-update')
  .description('PATCH a booking business (body: --json-file)')
  .argument('<businessId>', 'Booking business id')
  .requiredOption('--json-file <path>', 'JSON patch body per Graph bookingBusiness')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await updateBookingBusiness(auth.token, businessId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.displayName ?? ''}\t${r.data.id}`);
    }
  );

bookingsCommand
  .command('appointment-create')
  .description('Create an appointment (POST …/appointments; body: --json-file)')
  .argument('<businessId>', 'Booking business id')
  .requiredOption('--json-file <path>', 'JSON body per Graph bookingAppointment')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await createBookingAppointment(auth.token, businessId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

bookingsCommand
  .command('appointment-update')
  .description('PATCH an appointment')
  .argument('<businessId>', 'Booking business id')
  .argument('<appointmentId>', 'Appointment id')
  .requiredOption('--json-file <path>', 'JSON patch body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      appointmentId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await updateBookingAppointment(auth.token, businessId, appointmentId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

bookingsCommand
  .command('appointment-delete')
  .description('Delete an appointment')
  .argument('<businessId>', 'Booking business id')
  .argument('<appointmentId>', 'Appointment id')
  .option('--confirm', 'Required to perform delete', false)
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      appointmentId: string,
      opts: { confirm?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.error('Error: pass --confirm to delete an appointment.');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteBookingAppointment(auth.token, businessId, appointmentId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Deleted.');
    }
  );

bookingsCommand
  .command('appointment-cancel')
  .description('Cancel an appointment (POST …/cancel; optional JSON body)')
  .argument('<businessId>', 'Booking business id')
  .argument('<appointmentId>', 'Appointment id')
  .option('--json-file <path>', 'Optional POST body (e.g. cancellation message)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      appointmentId: string,
      opts: { jsonFile?: string; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown> | undefined;
      if (opts.jsonFile?.trim()) {
        body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      }
      const r = await cancelBookingAppointment(auth.token, businessId, appointmentId, body);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Cancelled.');
    }
  );

bookingsCommand
  .command('staff-availability')
  .description(
    'POST …/getStaffAvailability (Graph **application-only** — delegated tokens fail). Pass an app-only bearer via `--token`. Body: `--json-file` with staffIds + startDateTime + endDateTime (see Microsoft Graph).'
  )
  .argument('<businessId>', 'Booking business id')
  .requiredOption('--json-file <path>', 'JSON body per Graph (staffIds, startDateTime, endDateTime)')
  .option('--json', 'Print raw JSON response')
  .option('--token <token>', 'Graph access token (use **app-only**)')
  .option('--identity <name>', 'Not used for app-only; prefer `--token`')
  .action(async (businessId: string, opts: { jsonFile: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
    const r = await getBookingStaffAvailability(auth.token, businessId, body);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    const items = (r.data as any).staffAvailabilityItem ?? [];
    for (const item of items) {
      const staffId = item.staffId ?? '';
      const availCount = item.availabilityItems?.length ?? 0;
      console.log(`${staffId}\t${availCount} availability items`);
      for (const avail of item.availabilityItems ?? []) {
        const status = avail.status ?? '';
        const start = avail.startDateTime?.dateTime ?? '';
        const end = avail.endDateTime?.dateTime ?? '';
        console.log(`  ${status}\t${start}\t${end}`);
      }
    }
  });

function jsonFileAction(
  fn: (
    token: string,
    businessId: string,
    id: string,
    body: Record<string, unknown>
  ) => Promise<{ ok: boolean; data?: unknown; error?: { message?: string } }>
) {
  return async (
    businessId: string,
    id: string,
    opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
    cmd: Command
  ) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
    const r = await fn(auth.token, businessId, id, body);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${(r.data as any).id ?? ''}`);
  };
}

bookingsCommand
  .command('customer-create')
  .description('Create a customer (POST …/customers)')
  .argument('<businessId>', 'Booking business id')
  .requiredOption('--json-file <path>', 'JSON body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await createBookingCustomer(auth.token, businessId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

bookingsCommand
  .command('customer-update')
  .description('PATCH a customer')
  .argument('<businessId>', 'Booking business id')
  .argument('<customerId>', 'Customer id')
  .requiredOption('--json-file <path>', 'JSON patch body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    jsonFileAction(async (token, businessId, customerId, body) =>
      updateBookingCustomer(token, businessId, customerId, body)
    )
  );

bookingsCommand
  .command('customer-delete')
  .description('Delete a customer')
  .argument('<businessId>', 'Booking business id')
  .argument('<customerId>', 'Customer id')
  .option('--confirm', 'Required', false)
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      customerId: string,
      opts: { confirm?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.error('Error: pass --confirm to delete.');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteBookingCustomer(auth.token, businessId, customerId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Deleted.');
    }
  );

bookingsCommand
  .command('service-create')
  .description('Create a service (POST …/services)')
  .argument('<businessId>', 'Booking business id')
  .requiredOption('--json-file <path>', 'JSON body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await createBookingService(auth.token, businessId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

bookingsCommand
  .command('service-update')
  .description('PATCH a service')
  .argument('<businessId>', 'Booking business id')
  .argument('<serviceId>', 'Service id')
  .requiredOption('--json-file <path>', 'JSON patch body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    jsonFileAction(async (token, businessId, serviceId, body) =>
      updateBookingService(token, businessId, serviceId, body)
    )
  );

bookingsCommand
  .command('service-delete')
  .description('Delete a service')
  .argument('<businessId>', 'Booking business id')
  .argument('<serviceId>', 'Service id')
  .option('--confirm', 'Required', false)
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      serviceId: string,
      opts: { confirm?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.error('Error: pass --confirm to delete.');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteBookingService(auth.token, businessId, serviceId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Deleted.');
    }
  );

bookingsCommand
  .command('staff-create')
  .description('Create a staff member (POST …/staffMembers)')
  .argument('<businessId>', 'Booking business id')
  .requiredOption('--json-file <path>', 'JSON body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await createBookingStaffMember(auth.token, businessId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

bookingsCommand
  .command('staff-update')
  .description('PATCH a staff member')
  .argument('<businessId>', 'Booking business id')
  .argument('<staffId>', 'Staff id')
  .requiredOption('--json-file <path>', 'JSON patch body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    jsonFileAction(async (token, businessId, staffId, body) =>
      updateBookingStaffMember(token, businessId, staffId, body)
    )
  );

bookingsCommand
  .command('staff-delete')
  .description('Delete a staff member')
  .argument('<businessId>', 'Booking business id')
  .argument('<staffId>', 'Staff id')
  .option('--confirm', 'Required', false)
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      staffId: string,
      opts: { confirm?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.error('Error: pass --confirm to delete.');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteBookingStaffMember(auth.token, businessId, staffId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Deleted.');
    }
  );

bookingsCommand
  .command('custom-question-create')
  .description('Create a custom question (POST …/customQuestions)')
  .argument('<businessId>', 'Booking business id')
  .requiredOption('--json-file <path>', 'JSON body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      opts: { jsonFile: string; json?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const body = JSON.parse(await readFile(opts.jsonFile.trim(), 'utf-8')) as Record<string, unknown>;
      const r = await createBookingCustomQuestion(auth.token, businessId, body);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log(opts.json ? JSON.stringify(r.data, null, 2) : `${r.data.id ?? ''}`);
    }
  );

bookingsCommand
  .command('custom-question-update')
  .description('PATCH a custom question')
  .argument('<businessId>', 'Booking business id')
  .argument('<questionId>', 'Custom question id')
  .requiredOption('--json-file <path>', 'JSON patch body')
  .option('--json', 'Print response JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    jsonFileAction(async (token, businessId, questionId, body) =>
      updateBookingCustomQuestion(token, businessId, questionId, body)
    )
  );

bookingsCommand
  .command('custom-question-delete')
  .description('Delete a custom question')
  .argument('<businessId>', 'Booking business id')
  .argument('<questionId>', 'Custom question id')
  .option('--confirm', 'Required', false)
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      businessId: string,
      questionId: string,
      opts: { confirm?: boolean; token?: string; identity?: string },
      cmd: Command
    ) => {
      checkReadOnly(cmd);
      if (!opts.confirm) {
        console.error('Error: pass --confirm to delete.');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const r = await deleteBookingCustomQuestion(auth.token, businessId, questionId);
      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      console.log('Deleted.');
    }
  );
