import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  getBookingAppointment,
  getBookingBusiness,
  getBookingCustomer,
  getBookingService,
  getBookingStaffMember,
  listBookingAppointments,
  listBookingBusinesses,
  listBookingCalendarView,
  listBookingCurrencies,
  listBookingCustomQuestions,
  listBookingCustomers,
  listBookingServices,
  listBookingStaffMembers
} from '../lib/graph-bookings-client.js';

export const bookingsCommand = new Command('bookings').description(
  'Microsoft Bookings (Graph): businesses, business-get, currencies, appointments, customers, custom questions, services, service-get, staff, staff-get, calendar (delegated; see GRAPH_SCOPES.md)'
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
    async (
      businessId: string,
      serviceId: string,
      opts: { json?: boolean; token?: string; identity?: string }
    ) => {
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
  .action(
    async (
      businessId: string,
      staffId: string,
      opts: { json?: boolean; token?: string; identity?: string }
    ) => {
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
    }
  );

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
    async (
      businessId: string,
      customerId: string,
      opts: { json?: boolean; token?: string; identity?: string }
    ) => {
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
    async (
      businessId: string,
      appointmentId: string,
      opts: { json?: boolean; token?: string; identity?: string }
    ) => {
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
      const who =
        a.customerName ?? a.customers?.[0]?.name ?? a.customers?.[0]?.emailAddress ?? '';
      console.log(`${start}\t${end}\t${who}\t${a.id ?? ''}`);
    }
  );
