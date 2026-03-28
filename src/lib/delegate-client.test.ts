import { describe, expect, it } from 'bun:test';
import { getDelegates, addDelegate, updateDelegate, removeDelegate } from './delegate-client.js';

describe('delegate-client', () => {
  const token = 'test-token';

  it('getDelegates parses SOAP response properly', async () => {
    const fetchCalls: any[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input, init) => {
        fetchCalls.push({ input, init });
        const xml = `
          <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
            <s:Body>
              <m:GetDelegateResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
                <m:ResponseMessages>
                  <m:DelegateUserResponseMessageType ResponseClass="Success">
                    <m:DelegateUser>
                      <t:UserId>
                        <t:PrimarySmtpAddress>del@example.com</t:PrimarySmtpAddress>
                      </t:UserId>
                      <t:DelegatePermissions>
                        <t:CalendarFolderPermissionLevel>Editor</t:CalendarFolderPermissionLevel>
                      </t:DelegatePermissions>
                      <t:ViewPrivateItems>true</t:ViewPrivateItems>
                    </m:DelegateUser>
                  </m:DelegateUserResponseMessageType>
                </m:ResponseMessages>
                <m:DeliverMeetingRequests>DelegatesAndSendInformationToMe</m:DeliverMeetingRequests>
              </m:GetDelegateResponse>
            </s:Body>
          </s:Envelope>
        `;
        return new Response(xml, { status: 200, headers: { 'content-type': 'text/xml' } });
      }) as typeof fetch;

      const res = await getDelegates(token, 'me@example.com');
      expect(res.ok).toBe(true);
      expect(res.data?.length).toBe(1);
      expect(res.data?.[0].userId).toBe('del@example.com');
      expect(res.data?.[0].permissions.calendar).toBe('Editor');
      expect(res.data?.[0].viewPrivateItems).toBe(true);
      expect(res.data?.[0].deliverMeetingRequests).toBe('DelegatesAndSendInformationToMe');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('addDelegate generates correct SOAP body', async () => {
    const fetchCalls: any[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input, init) => {
        fetchCalls.push({ input, init });
        const xml = `<m:AddDelegateResponse ResponseClass="Success"></m:AddDelegateResponse>`;
        return new Response(xml, { status: 200, headers: { 'content-type': 'text/xml' } });
      }) as typeof fetch;

      await addDelegate({
        token,
        delegateEmail: 'del@example.com',
        permissions: { inbox: 'Reviewer' },
        deliverMeetingRequests: 'DelegatesAndSendInformationToMe'
      });
      
      const body = fetchCalls[0].init.body as string;
      expect(body).toContain('<m:DeliverMeetingRequests>DelegatesAndSendInformationToMe</m:DeliverMeetingRequests>');
      expect(body).toContain('<t:InboxFolderPermissionLevel>Reviewer</t:InboxFolderPermissionLevel>');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
