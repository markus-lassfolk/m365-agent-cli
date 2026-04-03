import { afterEach, describe, expect, it, mock } from 'bun:test';
import { mockResolveNamesResponse } from './mocks/responses.js';

const okUpdateResponse = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <m:UpdateItemResponse>
      <m:ResponseMessages>
        <m:UpdateItemResponseMessage ResponseClass="Success">
          <m:ResponseCode>NoError</m:ResponseCode>
          <m:Items>
            <t:CalendarItem>
              <t:ItemId Id="updated-id" ChangeKey="new-ck" />
            </t:CalendarItem>
          </m:Items>
        </m:UpdateItemResponseMessage>
      </m:ResponseMessages>
    </m:UpdateItemResponse>
  </soap:Body>
</soap:Envelope>`;

describe('ews-client safety and conflict behavior', () => {
  const originalFetch = globalThis.fetch;

  afterEach(() => {
    globalThis.fetch = originalFetch;
    mock.restore();
  });

  it('retries updateEvent with AlwaysOverwrite after conflict when ChangeKey is provided', async () => {
    const fetchCalls: string[] = [];
    let callCount = 0;

    globalThis.fetch = mock(async (_input: RequestInfo | URL, init?: RequestInit) => {
      const body = String(init?.body || '');
      fetchCalls.push(body);
      callCount += 1;

      if (callCount === 1) {
        return new Response(
          `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body>
    <m:UpdateItemResponse>
      <m:ResponseMessages>
        <m:UpdateItemResponseMessage ResponseClass="Error">
          <m:ResponseCode>ErrorIrresolvableConflict</m:ResponseCode>
          <m:MessageText>The change key passed in the request does not match.</m:MessageText>
        </m:UpdateItemResponseMessage>
      </m:ResponseMessages>
    </m:UpdateItemResponse>
  </soap:Body>
</soap:Envelope>`,
          { status: 200 }
        );
      }

      return new Response(okUpdateResponse, { status: 200 });
    }) as unknown as typeof fetch;

    const { updateEvent } = await import('../lib/ews-client.js');
    const result = await updateEvent({
      token: 'token',
      eventId: 'event-id',
      subject: 'Updated title',
      changeKey: 'client-ck'
    });

    expect(result.ok).toBe(true);
    expect(fetchCalls.length).toBe(2);
    expect(fetchCalls[0]).toContain('ConflictResolution="AutoResolve"');
    expect(fetchCalls[0]).toContain('<t:ItemId Id="event-id" ChangeKey="client-ck" />');
    expect(fetchCalls[1]).toContain('ConflictResolution="AlwaysOverwrite"');
    expect(fetchCalls[1]).toContain('<t:ItemId Id="event-id" />');
  });

  it('sanitizes EWS QueryString control syntax in getEmails search', async () => {
    const fetchCalls: string[] = [];

    globalThis.fetch = mock(async (_input: RequestInfo | URL, init?: RequestInit) => {
      const body = String(init?.body || '');
      fetchCalls.push(body);
      return new Response(
        `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body>
    <m:FindItemResponse>
      <m:ResponseMessages>
        <m:FindItemResponseMessage ResponseClass="Success">
          <m:ResponseCode>NoError</m:ResponseCode>
          <m:RootFolder IncludesLastItemInRange="true" TotalItemsInView="0" IndexedPagingOffset="0" />
        </m:FindItemResponseMessage>
      </m:ResponseMessages>
    </m:FindItemResponse>
  </soap:Body>
</soap:Envelope>`,
        { status: 200 }
      );
    }) as unknown as typeof fetch;

    const { getEmails } = await import('../lib/ews-client.js');
    const query = 'urgent OR from:bob@example.com AND "project x"';
    const result = await getEmails({ token: 'token', search: query });

    expect(result.ok).toBe(true);
    expect(fetchCalls.length).toBe(1);
    expect(fetchCalls[0]).toContain(
      '<m:QueryString>urgent OR from:bob@example.com AND &quot;project x&quot;</m:QueryString>'
    );
  });

  it('returns explicit error when getOwaUserInfo fails instead of silent fallback', async () => {
    globalThis.fetch = mock(async () => {
      return new Response('gateway timeout', { status: 504 });
    }) as unknown as typeof fetch;

    const { getOwaUserInfo } = await import('../lib/ews-client.js');
    const result = await getOwaUserInfo('token');

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe('EWS_ERROR');
    expect(result.error?.message).toContain('Failed to resolve OWA user info');
  });

  it('getOwaUserInfo embeds process.env.EWS_USERNAME at call time in ResolveNames', async () => {
    const prev = process.env.EWS_USERNAME;
    const bodies: string[] = [];
    globalThis.fetch = mock(async (_input: RequestInfo | URL, init?: RequestInit) => {
      bodies.push(String(init?.body || ''));
      return new Response(mockResolveNamesResponse, { status: 200, headers: { 'content-type': 'text/xml' } });
    }) as unknown as typeof fetch;

    try {
      delete process.env.EWS_USERNAME;
      const { getOwaUserInfo } = await import('../lib/ews-client.js');
      await getOwaUserInfo('tok');
      expect(bodies[0]).toContain('<m:UnresolvedEntry></m:UnresolvedEntry>');

      process.env.EWS_USERNAME = 'alice@contoso.com';
      await getOwaUserInfo('tok');
      expect(bodies[1]).toContain('alice@contoso.com');
      expect(bodies[1]).toContain('<m:UnresolvedEntry>');
    } finally {
      if (prev === undefined) {
        delete process.env.EWS_USERNAME;
      } else {
        process.env.EWS_USERNAME = prev;
      }
    }
  });

  it('parses TimeZone correctly from CalendarItem StartTimeZone and EndTimeZone', async () => {
    globalThis.fetch = mock(async () => {
      return new Response(
        `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <m:GetItemResponse>
      <m:ResponseMessages>
        <m:GetItemResponseMessage ResponseClass="Success">
          <m:ResponseCode>NoError</m:ResponseCode>
          <m:Items>
            <t:CalendarItem>
              <t:ItemId Id="event-id" ChangeKey="ck" />
              <t:Subject>Timezone Test Event</t:Subject>
              <t:Start>2026-03-30T10:00:00Z</t:Start>
              <t:End>2026-03-30T11:00:00Z</t:End>
              <t:StartTimeZone Id="Pacific Standard Time" />
              <t:EndTimeZone Id="Pacific Standard Time" />
              <t:IsAllDayEvent>false</t:IsAllDayEvent>
              <t:IsCancelled>false</t:IsCancelled>
              <t:Organizer><t:Mailbox><t:Name>Bob</t:Name><t:EmailAddress>bob@example.com</t:EmailAddress></t:Mailbox></t:Organizer>
            </t:CalendarItem>
          </m:Items>
        </m:GetItemResponseMessage>
      </m:ResponseMessages>
    </m:GetItemResponse>
  </soap:Body>
</soap:Envelope>`,
        { status: 200 }
      );
    }) as unknown as typeof fetch;

    const { getCalendarEvent } = await import('../lib/ews-client.js');
    const result = await getCalendarEvent('token', 'event-id');

    expect(result.ok).toBe(true);
    expect(result.data?.Start.TimeZone).toBe('Pacific Standard Time');
    expect(result.data?.End.TimeZone).toBe('Pacific Standard Time');
  });

  it('replyToEmail sends ReferenceItemId with ChangeKey after GetItem', async () => {
    const fetchCalls: string[] = [];
    let callCount = 0;

    globalThis.fetch = mock(async (_input: RequestInfo | URL, init?: RequestInit) => {
      const body = String(init?.body || '');
      fetchCalls.push(body);
      callCount += 1;
      if (callCount === 1) {
        return new Response(
          `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:GetItemResponse>
      <m:Items>
        <t:Message>
          <t:ItemId Id="msg-1" ChangeKey="ck-from-get" />
          <t:Subject>Subj</t:Subject>
        </t:Message>
      </m:Items>
    </m:GetItemResponse>
  </soap:Body>
</soap:Envelope>`,
          { status: 200 }
        );
      }
      return new Response(
        `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body>
    <m:ResponseCode>NoError</m:ResponseCode>
  </soap:Body>
</soap:Envelope>`,
        { status: 200 }
      );
    }) as unknown as typeof fetch;

    const { replyToEmail } = await import('../lib/ews-client.js');
    const result = await replyToEmail('token', 'msg-1', 'Thanks', false, false, undefined);

    expect(result.ok).toBe(true);
    expect(fetchCalls.length).toBe(2);
    expect(fetchCalls[0]).toContain('<m:GetItem>');
    expect(fetchCalls[1]).toContain('ReferenceItemId');
    expect(fetchCalls[1]).toContain('ChangeKey="ck-from-get"');
    expect(fetchCalls[1]).toContain('Id="msg-1"');
  });

  it('replyToEmailDraft sends ReferenceItemId with ChangeKey after GetItem', async () => {
    const fetchCalls: string[] = [];
    let callCount = 0;

    globalThis.fetch = mock(async (_input: RequestInfo | URL, init?: RequestInit) => {
      const body = String(init?.body || '');
      fetchCalls.push(body);
      callCount += 1;
      if (callCount === 1) {
        return new Response(
          `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:GetItemResponse>
      <m:Items>
        <t:Message>
          <t:ItemId Id="msg-2" ChangeKey="ck-draft" />
          <t:Subject>Subj</t:Subject>
        </t:Message>
      </m:Items>
    </m:GetItemResponse>
  </soap:Body>
</soap:Envelope>`,
          { status: 200 }
        );
      }
      return new Response(
        `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <m:ResponseCode>NoError</m:ResponseCode>
    <m:CreateItemResponse>
      <m:Items>
        <t:Message>
          <t:ItemId Id="reply-draft-x" ChangeKey="rck" />
        </t:Message>
      </m:Items>
    </m:CreateItemResponse>
  </soap:Body>
</soap:Envelope>`,
        { status: 200 }
      );
    }) as unknown as typeof fetch;

    const { replyToEmailDraft } = await import('../lib/ews-client.js');
    const result = await replyToEmailDraft('token', 'msg-2', 'Draft reply', false, false, undefined);

    expect(result.ok).toBe(true);
    expect(result.data?.draftId).toBe('reply-draft-x');
    expect(fetchCalls.length).toBe(2);
    expect(fetchCalls[0]).toContain('<m:GetItem>');
    expect(fetchCalls[1]).toContain('ChangeKey="ck-draft"');
    expect(fetchCalls[1]).toContain('MessageDisposition="SaveOnly"');
  });
});
