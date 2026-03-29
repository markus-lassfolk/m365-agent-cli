import { serve } from 'bun';

export function startWebhookServer(port: number = 3000) {
  console.log(`Starting webhook receiver on http://localhost:${port}/webhooks/m365-agent-cli`);
  serve({
    port,
    async fetch(req) {
      const url = new URL(req.url);
      if (url.pathname === '/webhooks/m365-agent-cli') {
        // Microsoft Graph sends validationToken as a query param on POST (not GET)
        const validationToken = url.searchParams.get('validationToken');
        if (validationToken) {
          console.log(`[${new Date().toISOString()}] Received validation token request. Replaying token...`);
          return new Response(validationToken, {
            status: 200,
            headers: { 'Content-Type': 'text/plain' }
          });
        }
        if (req.method === 'POST') {
          try {
            const body = await req.json();

            // Validate clientState if configured
            const expectedClientState = process.env.GRAPH_CLIENT_STATE;
            const notifications = Array.isArray((body as any).value) ? (body as any).value : null;
            if (expectedClientState) {
              const allClientStatesValid =
                !!notifications &&
                notifications.length > 0 &&
                notifications.every((n: any) => n && n.clientState === expectedClientState);
              if (!allClientStatesValid) {
                console.warn(
                  `[${new Date().toISOString()}] Received Graph notification with invalid or missing clientState.`
                );
                return new Response('Invalid clientState', { status: 401 });
              }
            }

            console.log(`[${new Date().toISOString()}] Received Graph notification:`);
            console.log(JSON.stringify(body, null, 2));
            return new Response('Accepted', { status: 202 });
          } catch (err) {
            console.error('Error parsing notification body:', err);
            return new Response('Bad Request', { status: 400 });
          }
        }
      }
      return new Response('Not Found', { status: 404 });
    }
  });
}
