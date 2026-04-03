import {
  callGraph,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';

export interface BookingBusiness {
  id: string;
  displayName?: string;
  businessType?: string;
  webSiteUrl?: string;
}

export interface BookingAppointment {
  id?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  serviceNotes?: string;
  customerName?: string;
  customerEmailAddress?: string;
  customers?: Array<{ name?: string; emailAddress?: string }>;
}

export interface BookingService {
  id?: string;
  displayName?: string;
  defaultDuration?: string;
}

export interface BookingStaffMember {
  id?: string;
  displayName?: string;
  emailAddress?: string;
  role?: string;
}

export interface BookingCustomer {
  id?: string;
  displayName?: string;
  emailAddress?: string;
}

export interface BookingCustomQuestion {
  id?: string;
  displayName?: string;
  answerInputType?: string;
  answerOptions?: string[];
}

export async function listBookingBusinesses(token: string): Promise<GraphResponse<BookingBusiness[]>> {
  try {
    const r = await callGraph<{ value: BookingBusiness[] }>(token, '/solutions/bookingBusinesses');
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list booking businesses', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list booking businesses');
  }
}

export async function getBookingBusiness(
  token: string,
  businessId: string
): Promise<GraphResponse<BookingBusiness>> {
  try {
    const r = await callGraph<BookingBusiness>(
      token,
      `/solutions/bookingBusinesses/${encodeURIComponent(businessId)}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get booking business', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get booking business');
  }
}

export async function listBookingAppointments(
  token: string,
  businessId: string
): Promise<GraphResponse<BookingAppointment[]>> {
  try {
    const r = await callGraph<{ value: BookingAppointment[] }>(
      token,
      `/solutions/bookingBusinesses/${encodeURIComponent(businessId)}/appointments`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list appointments', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list appointments');
  }
}

export async function listBookingServices(
  token: string,
  businessId: string
): Promise<GraphResponse<BookingService[]>> {
  try {
    const r = await callGraph<{ value: BookingService[] }>(
      token,
      `/solutions/bookingBusinesses/${encodeURIComponent(businessId)}/services`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list booking services', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list booking services');
  }
}

export async function getBookingService(
  token: string,
  businessId: string,
  serviceId: string
): Promise<GraphResponse<BookingService>> {
  try {
    const b = encodeURIComponent(businessId);
    const s = encodeURIComponent(serviceId);
    const r = await callGraph<BookingService>(token, `/solutions/bookingBusinesses/${b}/services/${s}`);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get booking service', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get booking service');
  }
}

export async function listBookingStaffMembers(
  token: string,
  businessId: string
): Promise<GraphResponse<BookingStaffMember[]>> {
  try {
    const r = await callGraph<{ value: BookingStaffMember[] }>(
      token,
      `/solutions/bookingBusinesses/${encodeURIComponent(businessId)}/staffMembers`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list staff members', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list staff members');
  }
}

export async function getBookingStaffMember(
  token: string,
  businessId: string,
  staffId: string
): Promise<GraphResponse<BookingStaffMember>> {
  try {
    const b = encodeURIComponent(businessId);
    const s = encodeURIComponent(staffId);
    const r = await callGraph<BookingStaffMember>(
      token,
      `/solutions/bookingBusinesses/${b}/staffMembers/${s}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get staff member', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get staff member');
  }
}

/** `start` and `end` must be ISO 8601 date-times (UTC recommended), e.g. 2026-04-01T00:00:00Z */
export async function listBookingCalendarView(
  token: string,
  businessId: string,
  start: string,
  end: string
): Promise<GraphResponse<BookingAppointment[]>> {
  try {
    const q = `?start=${encodeURIComponent(start.trim())}&end=${encodeURIComponent(end.trim())}`;
    const r = await callGraph<{ value: BookingAppointment[] }>(
      token,
      `/solutions/bookingBusinesses/${encodeURIComponent(businessId)}/calendarView${q}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list calendar view', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list calendar view');
  }
}

export async function listBookingCustomers(
  token: string,
  businessId: string
): Promise<GraphResponse<BookingCustomer[]>> {
  try {
    const r = await callGraph<{ value: BookingCustomer[] }>(
      token,
      `/solutions/bookingBusinesses/${encodeURIComponent(businessId)}/customers`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list customers', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list customers');
  }
}

export async function getBookingCustomer(
  token: string,
  businessId: string,
  customerId: string
): Promise<GraphResponse<BookingCustomer>> {
  try {
    const r = await callGraph<BookingCustomer>(
      token,
      `/solutions/bookingBusinesses/${encodeURIComponent(businessId)}/customers/${encodeURIComponent(customerId)}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get customer', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get customer');
  }
}

export async function getBookingAppointment(
  token: string,
  businessId: string,
  appointmentId: string
): Promise<GraphResponse<BookingAppointment>> {
  try {
    const r = await callGraph<BookingAppointment>(
      token,
      `/solutions/bookingBusinesses/${encodeURIComponent(businessId)}/appointments/${encodeURIComponent(appointmentId)}`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get appointment', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get appointment');
  }
}

export async function listBookingCustomQuestions(
  token: string,
  businessId: string
): Promise<GraphResponse<BookingCustomQuestion[]>> {
  try {
    const r = await callGraph<{ value: BookingCustomQuestion[] }>(
      token,
      `/solutions/bookingBusinesses/${encodeURIComponent(businessId)}/customQuestions`
    );
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list custom questions', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list custom questions');
  }
}

export interface BookingCurrency {
  id?: string;
  symbol?: string;
}

/** Tenant-wide currency catalog for Bookings (not scoped to one business). */
export async function listBookingCurrencies(token: string): Promise<GraphResponse<BookingCurrency[]>> {
  try {
    const r = await callGraph<{ value: BookingCurrency[] }>(token, '/solutions/bookingCurrencies');
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to list booking currencies', r.error?.code, r.error?.status);
    }
    return graphResult(r.data.value ?? []);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to list booking currencies');
  }
}
