import { callGraph, GraphApiError, type GraphResponse, graphError, graphResult } from './graph-client.js';

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

export async function getBookingBusiness(token: string, businessId: string): Promise<GraphResponse<BookingBusiness>> {
  try {
    const r = await callGraph<BookingBusiness>(token, `/solutions/bookingBusinesses/${encodeURIComponent(businessId)}`);
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

export async function listBookingServices(token: string, businessId: string): Promise<GraphResponse<BookingService[]>> {
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
    const r = await callGraph<BookingStaffMember>(token, `/solutions/bookingBusinesses/${b}/staffMembers/${s}`);
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

function businessPath(businessId: string): string {
  return `/solutions/bookingBusinesses/${encodeURIComponent(businessId.trim())}`;
}

export async function createBookingAppointment(
  token: string,
  businessId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<BookingAppointment>> {
  try {
    const r = await callGraph<BookingAppointment>(token, `${businessPath(businessId)}/appointments`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create appointment', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create appointment');
  }
}

export async function updateBookingAppointment(
  token: string,
  businessId: string,
  appointmentId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<BookingAppointment>> {
  try {
    const path = `${businessPath(businessId)}/appointments/${encodeURIComponent(appointmentId)}`;
    const r = await callGraph<BookingAppointment>(token, path, {
      method: 'PATCH',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to update appointment', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update appointment');
  }
}

export async function deleteBookingAppointment(
  token: string,
  businessId: string,
  appointmentId: string
): Promise<GraphResponse<void>> {
  try {
    const path = `${businessPath(businessId)}/appointments/${encodeURIComponent(appointmentId)}`;
    const r = await callGraph<void>(token, path, { method: 'DELETE' }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete appointment', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete appointment');
  }
}

export async function cancelBookingAppointment(
  token: string,
  businessId: string,
  appointmentId: string,
  body?: Record<string, unknown>
): Promise<GraphResponse<void>> {
  try {
    const path = `${businessPath(businessId)}/appointments/${encodeURIComponent(appointmentId)}/cancel`;
    const r = await callGraph<void>(token, path, { method: 'POST', body: JSON.stringify(body ?? {}) }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to cancel appointment', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to cancel appointment');
  }
}

export async function createBookingCustomer(
  token: string,
  businessId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<BookingCustomer>> {
  try {
    const r = await callGraph<BookingCustomer>(token, `${businessPath(businessId)}/customers`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create customer', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create customer');
  }
}

export async function updateBookingCustomer(
  token: string,
  businessId: string,
  customerId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<BookingCustomer>> {
  try {
    const path = `${businessPath(businessId)}/customers/${encodeURIComponent(customerId)}`;
    const r = await callGraph<BookingCustomer>(token, path, {
      method: 'PATCH',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to update customer', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update customer');
  }
}

export async function deleteBookingCustomer(
  token: string,
  businessId: string,
  customerId: string
): Promise<GraphResponse<void>> {
  try {
    const path = `${businessPath(businessId)}/customers/${encodeURIComponent(customerId)}`;
    const r = await callGraph<void>(token, path, { method: 'DELETE' }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete customer', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete customer');
  }
}

export async function createBookingService(
  token: string,
  businessId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<BookingService>> {
  try {
    const r = await callGraph<BookingService>(token, `${businessPath(businessId)}/services`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create service', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create service');
  }
}

export async function updateBookingService(
  token: string,
  businessId: string,
  serviceId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<BookingService>> {
  try {
    const path = `${businessPath(businessId)}/services/${encodeURIComponent(serviceId)}`;
    const r = await callGraph<BookingService>(token, path, {
      method: 'PATCH',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to update service', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update service');
  }
}

export async function deleteBookingService(
  token: string,
  businessId: string,
  serviceId: string
): Promise<GraphResponse<void>> {
  try {
    const path = `${businessPath(businessId)}/services/${encodeURIComponent(serviceId)}`;
    const r = await callGraph<void>(token, path, { method: 'DELETE' }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete service', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete service');
  }
}

export async function createBookingStaffMember(
  token: string,
  businessId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<BookingStaffMember>> {
  try {
    const r = await callGraph<BookingStaffMember>(token, `${businessPath(businessId)}/staffMembers`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create staff member', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create staff member');
  }
}

export async function updateBookingStaffMember(
  token: string,
  businessId: string,
  staffId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<BookingStaffMember>> {
  try {
    const path = `${businessPath(businessId)}/staffMembers/${encodeURIComponent(staffId)}`;
    const r = await callGraph<BookingStaffMember>(token, path, {
      method: 'PATCH',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to update staff member', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update staff member');
  }
}

export async function deleteBookingStaffMember(
  token: string,
  businessId: string,
  staffId: string
): Promise<GraphResponse<void>> {
  try {
    const path = `${businessPath(businessId)}/staffMembers/${encodeURIComponent(staffId)}`;
    const r = await callGraph<void>(token, path, { method: 'DELETE' }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete staff member', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete staff member');
  }
}

export async function createBookingCustomQuestion(
  token: string,
  businessId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<BookingCustomQuestion>> {
  try {
    const r = await callGraph<BookingCustomQuestion>(token, `${businessPath(businessId)}/customQuestions`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create custom question', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create custom question');
  }
}

export async function updateBookingCustomQuestion(
  token: string,
  businessId: string,
  questionId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<BookingCustomQuestion>> {
  try {
    const path = `${businessPath(businessId)}/customQuestions/${encodeURIComponent(questionId)}`;
    const r = await callGraph<BookingCustomQuestion>(token, path, {
      method: 'PATCH',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to update custom question', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update custom question');
  }
}

export async function deleteBookingCustomQuestion(
  token: string,
  businessId: string,
  questionId: string
): Promise<GraphResponse<void>> {
  try {
    const path = `${businessPath(businessId)}/customQuestions/${encodeURIComponent(questionId)}`;
    const r = await callGraph<void>(token, path, { method: 'DELETE' }, false);
    if (!r.ok) {
      return graphError(r.error?.message || 'Failed to delete custom question', r.error?.code, r.error?.status);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete custom question');
  }
}

export async function updateBookingBusiness(
  token: string,
  businessId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<BookingBusiness>> {
  try {
    const r = await callGraph<BookingBusiness>(token, businessPath(businessId), {
      method: 'PATCH',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to update booking business', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update booking business');
  }
}

/** Microsoft Graph documents **application-only** for this API (no delegated). Use an app-only access token (`--token`). */
export async function getBookingStaffAvailability(
  token: string,
  businessId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<unknown>> {
  try {
    const r = await callGraph<unknown>(token, `${businessPath(businessId)}/getStaffAvailability`, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get staff availability', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get staff availability');
  }
}
