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

  it('getEmails combines isRead and flagStatus into And restriction', async () => {
    const bodies: string[] = [];
    const emptyFindItem = `<?xml version="1.0" encoding="utf-8"?>
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
</soap:Envelope>`;

    globalThis.fetch = mock(async (_input: RequestInfo | URL, init?: RequestInit) => {
      bodies.push(String(init?.body || ''));
      return new Response(emptyFindItem, { status: 200 });
    }) as unknown as typeof fetch;

    const { getEmails } = await import('../lib/ews-client.js');
    const result = await getEmails({ token: 'token', isRead: false, flagStatus: 'Flagged' });

    expect(result.ok).toBe(true);
    expect(bodies[0]).toContain('<t:And>');
    expect(bodies[0]).toContain('message:IsRead');
    expect(bodies[0]).toContain('item:Flag/FlagStatus');
    expect(bodies[0]).toContain('Value="Flagged"');
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

  it('getCalendarEvents pages a truncated CalendarView by advancing StartDate and de-dupes the boundary item', async () => {
    const startDates: string[] = [];
    let callCount = 0;
    const calItem = (id: string, start: string, end: string) =>
      `<t:CalendarItem><t:ItemId Id="${id}" ChangeKey="ck" /><t:Subject>E ${id}</t:Subject><t:Start>${start}</t:Start><t:End>${end}</t:End><t:IsAllDayEvent>false</t:IsAllDayEvent><t:IsCancelled>false</t:IsCancelled></t:CalendarItem>`;
    const findResp = (includesLast: boolean, items: string) =>
      `<?xml version="1.0"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"><soap:Body><m:ResponseCode>NoError</m:ResponseCode><m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder TotalItemsInView="3" IncludesLastItemInRange="${includesLast}"><t:Items>${items}</t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse></soap:Body></soap:Envelope>`;

    globalThis.fetch = mock(async (_input: RequestInfo | URL, init?: RequestInit) => {
      const body = String(init?.body || '');
      const m = body.match(/StartDate="([^"]+)"/);
      if (m) startDates.push(m[1]);
      callCount += 1;
      if (callCount === 1) {
        // First page truncated: two items, last starts at 2026-01-02T09:00:00Z.
        return new Response(
          findResp(
            false,
            calItem('a', '2026-01-01T09:00:00Z', '2026-01-01T10:00:00Z') +
              calItem('b', '2026-01-02T09:00:00Z', '2026-01-02T10:00:00Z')
          ),
          { status: 200 }
        );
      }
      // Second page (StartDate advanced): boundary item 'b' repeats (de-duped) + new item 'c'; end of range.
      return new Response(
        findResp(
          true,
          calItem('b', '2026-01-02T09:00:00Z', '2026-01-02T10:00:00Z') +
            calItem('c', '2026-01-03T09:00:00Z', '2026-01-03T10:00:00Z')
        ),
        { status: 200 }
      );
    }) as unknown as typeof fetch;

    const { getCalendarEvents } = await import('../lib/ews-client.js');
    const result = await getCalendarEvents('token', '2026-01-01T00:00:00Z', '2026-01-31T00:00:00Z');

    expect(result.ok).toBe(true);
    // Three unique events, boundary 'b' counted once.
    expect(result.data?.length).toBe(3);
    expect(result.data?.map((e) => e.Id).sort()).toEqual(['a', 'b', 'c']);
    // Two requests: first at the window start, second advanced to the last item's start.
    expect(callCount).toBe(2);
    expect(startDates[0]).toBe('2026-01-01T00:00:00Z');
    expect(startDates[1]).toBe('2026-01-02T09:00:00Z');
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

  it('replyToEmail includes CcRecipients and BccRecipients when provided', async () => {
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
    const result = await replyToEmail('token', 'msg-1', 'Thanks', true, false, undefined, {
      cc: ['cc1@contoso.com'],
      bcc: ['bcc1@contoso.com', 'bcc2@contoso.com']
    });

    expect(result.ok).toBe(true);
    expect(fetchCalls[1]).toContain('<t:ReplyAllToItem>');
    expect(fetchCalls[1]).toContain('<t:CcRecipients>');
    expect(fetchCalls[1]).toContain('cc1@contoso.com');
    expect(fetchCalls[1]).toContain('<t:BccRecipients>');
    expect(fetchCalls[1]).toContain('bcc1@contoso.com');
    expect(fetchCalls[1]).toContain('bcc2@contoso.com');

    // The EWS Types.xsd ReplyToItem sequence requires Cc/BccRecipients BEFORE ReferenceItemId,
    // and NewBodyContent AFTER it. Out-of-order elements => ErrorSchemaValidation (reply never sends).
    const xml = fetchCalls[1];
    const iCc = xml.indexOf('<t:CcRecipients>');
    const iBcc = xml.indexOf('<t:BccRecipients>');
    const iRef = xml.indexOf('ReferenceItemId');
    const iBody = xml.indexOf('<t:NewBodyContent');
    expect(iCc).toBeGreaterThan(-1);
    expect(iBcc).toBeGreaterThan(-1);
    expect(iRef).toBeGreaterThan(-1);
    expect(iBody).toBeGreaterThan(-1);
    expect(iCc).toBeLessThan(iRef);
    expect(iBcc).toBeLessThan(iRef);
    expect(iRef).toBeLessThan(iBody);
  });

  it('sendEmail with an attachment threads the CreateAttachment change key into SendItem', async () => {
    const bodies: string[] = [];
    let call = 0;
    const soap = (inner: string) =>
      `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body><m:ResponseCode>NoError</m:ResponseCode>${inner}</soap:Body>
</soap:Envelope>`;

    globalThis.fetch = mock(async (_input: RequestInfo | URL, init?: RequestInit) => {
      bodies.push(String(init?.body || ''));
      call += 1;
      if (call === 1) {
        // createDraft (CreateItem SaveOnly) — returns the draft id and its initial change key.
        return new Response(
          soap(
            '<m:CreateItemResponse><m:Items><t:Message><t:ItemId Id="draft-1" ChangeKey="ck0" /></t:Message></m:Items></m:CreateItemResponse>'
          ),
          { status: 200 }
        );
      }
      if (call === 2) {
        // CreateAttachment — returns the new root change key as attributes on <t:AttachmentId>.
        return new Response(
          soap(
            '<m:CreateAttachmentResponse><m:Attachments><t:FileAttachment><t:AttachmentId Id="att-1" RootItemId="draft-1" RootItemChangeKey="ck1-updated" /></t:FileAttachment></m:Attachments></m:CreateAttachmentResponse>'
          ),
          { status: 200 }
        );
      }
      // SendItem
      return new Response(soap('<m:SendItemResponse />'), { status: 200 });
    }) as unknown as typeof fetch;

    const { sendEmail } = await import('../lib/ews-client.js');
    const result = await sendEmail('token', {
      to: ['a@b.com'],
      subject: 'S',
      body: 'B',
      attachments: [{ name: 'f.txt', contentType: 'text/plain', contentBytes: Buffer.from('x').toString('base64') }]
    });

    expect(result.ok).toBe(true);
    expect(call).toBe(3);
    // The final SendItem must carry the change key returned by CreateAttachment, not the stale ck0.
    expect(bodies[2]).toContain('ck1-updated');
    expect(bodies[2]).not.toContain('ck0');
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

describe('ewsAvailabilityTimeToUtcMs', () => {
  it('parses zone-less EWS availability times as UTC (not host-local)', async () => {
    const { ewsAvailabilityTimeToUtcMs } = await import('../lib/ews-client.js');
    // A zone-less value must equal the same instant with an explicit Z.
    expect(ewsAvailabilityTimeToUtcMs('2026-07-12T10:00:00')).toBe(Date.UTC(2026, 6, 12, 10, 0, 0));
    expect(ewsAvailabilityTimeToUtcMs('2026-07-12T10:00:00.000')).toBe(Date.UTC(2026, 6, 12, 10, 0, 0));
    // A value that already carries Z or an offset is respected.
    expect(ewsAvailabilityTimeToUtcMs('2026-07-12T10:00:00Z')).toBe(Date.UTC(2026, 6, 12, 10, 0, 0));
    expect(ewsAvailabilityTimeToUtcMs('2026-07-12T12:00:00+02:00')).toBe(Date.UTC(2026, 6, 12, 10, 0, 0));
    expect(Number.isNaN(ewsAvailabilityTimeToUtcMs(''))).toBe(true);
    expect(Number.isNaN(ewsAvailabilityTimeToUtcMs(undefined))).toBe(true);
  });

  it('getAutoReplyRule finds the template rule regardless of XML namespace prefix', async () => {
    // Response uses an `n0:` types prefix instead of `t:` — the old hardcoded /<t:Rule>/ regex
    // would miss it (and later create a duplicate); the prefix-agnostic extractors must find it.
    globalThis.fetch = mock(async () => {
      return new Response(
        `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:n0="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <m:GetInboxRulesResponse>
      <m:ResponseCode>NoError</m:ResponseCode>
      <m:InboxRules>
        <n0:Rule>
          <n0:DisplayName>AutoReplyTemplate</n0:DisplayName>
          <n0:IsEnabled>true</n0:IsEnabled>
        </n0:Rule>
      </m:InboxRules>
    </m:GetInboxRulesResponse>
  </soap:Body>
</soap:Envelope>`,
        { status: 200 }
      );
    }) as unknown as typeof fetch;

    const { getAutoReplyRule } = await import('../lib/ews-client.js');
    const result = await getAutoReplyRule('token', 'user@example.com');
    expect(result.ok).toBe(true);
    expect(result.data).not.toBeNull();
    expect(result.data?.enabled).toBe(true);
  });

  it('forwardEmail emits ToRecipients/Cc/Bcc before ReferenceItemId and NewBodyContent after (Types.xsd order)', async () => {
    const fetchCalls: string[] = [];
    let callCount = 0;
    globalThis.fetch = mock(async (_input: RequestInfo | URL, init?: RequestInit) => {
      fetchCalls.push(String(init?.body || ''));
      callCount += 1;
      if (callCount === 1) {
        return new Response(
          `<?xml version="1.0"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"><soap:Body><m:ResponseCode>NoError</m:ResponseCode><m:GetItemResponse><m:Items><t:Message><t:ItemId Id="msg-9" ChangeKey="ck-9" /></t:Message></m:Items></m:GetItemResponse></soap:Body></soap:Envelope>`,
          { status: 200 }
        );
      }
      return new Response(
        `<?xml version="1.0"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"><soap:Body><m:ResponseCode>NoError</m:ResponseCode></soap:Body></soap:Envelope>`,
        { status: 200 }
      );
    }) as unknown as typeof fetch;

    const { forwardEmail } = await import('../lib/ews-client.js');
    const result = await forwardEmail('token', 'msg-9', ['to@contoso.com'], 'FYI', undefined, {
      cc: ['cc@contoso.com'],
      bcc: ['bcc@contoso.com']
    });
    expect(result.ok).toBe(true);
    const xml = fetchCalls[1];
    const iTo = xml.indexOf('<t:ToRecipients>');
    const iCc = xml.indexOf('<t:CcRecipients>');
    const iBcc = xml.indexOf('<t:BccRecipients>');
    const iRef = xml.indexOf('ReferenceItemId');
    const iBody = xml.indexOf('<t:NewBodyContent');
    for (const idx of [iTo, iCc, iBcc, iRef, iBody]) expect(idx).toBeGreaterThan(-1);
    expect(iTo).toBeLessThan(iCc);
    expect(iCc).toBeLessThan(iBcc);
    expect(iBcc).toBeLessThan(iRef);
    expect(iRef).toBeLessThan(iBody);
  });

  it('respondToEvent emits Body before ReferenceItemId (AcceptItem Types.xsd order)', async () => {
    const fetchCalls: string[] = [];
    let callCount = 0;
    globalThis.fetch = mock(async (_input: RequestInfo | URL, init?: RequestInit) => {
      fetchCalls.push(String(init?.body || ''));
      callCount += 1;
      if (callCount === 1) {
        return new Response(
          `<?xml version="1.0"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"><soap:Body><m:ResponseCode>NoError</m:ResponseCode><m:GetItemResponse><m:Items><t:CalendarItem><t:ItemId Id="cal-1" ChangeKey="ck-c" /></t:CalendarItem></m:Items></m:GetItemResponse></soap:Body></soap:Envelope>`,
          { status: 200 }
        );
      }
      return new Response(
        `<?xml version="1.0"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"><soap:Body><m:ResponseCode>NoError</m:ResponseCode></soap:Body></soap:Envelope>`,
        { status: 200 }
      );
    }) as unknown as typeof fetch;

    const { respondToEvent } = await import('../lib/ews-client.js');
    const result = await respondToEvent({
      token: 'token',
      eventId: 'cal-1',
      response: 'accept',
      comment: 'See you there',
      sendResponse: true
    });
    expect(result.ok).toBe(true);
    const xml = fetchCalls[1];
    const iBody = xml.indexOf('<t:Body');
    const iRef = xml.indexOf('ReferenceItemId');
    expect(iBody).toBeGreaterThan(-1);
    expect(iRef).toBeGreaterThan(-1);
    expect(iBody).toBeLessThan(iRef);
    expect(xml).toContain('<t:AcceptItem>');
  });
});
