import { createServer, type IncomingMessage, type ServerResponse } from 'node:http';

async function readJsonBody(req: IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = [];
  for await (const chunk of req) {
    chunks.push(chunk as Buffer);
  }
  const raw = Buffer.concat(chunks).toString('utf8');
  if (!raw.trim()) return undefined;
  return JSON.parse(raw) as unknown;
}

function sendResponse(res: ServerResponse, status: number, body: string, contentType = 'text/plain'): void {
  res.writeHead(status, { 'Content-Type': contentType });
  res.end(body);
}

/**
 * Minimal webhook receiver for Graph subscription notifications.
 * Uses Node `http` so the CLI can load under Node/tsx without the Bun runtime.
 */
export function startWebhookServer(port: number = 3000): void {
  console.log(`Starting webhook receiver on http://localhost:${port}/webhooks/m365-agent-cli`);

  const server = createServer(async (req, res) => {
    try {
      const host = req.headers.host ?? `localhost:${port}`;
      const url = new URL(req.url ?? '/', `http://${host}`);

      if (url.pathname !== '/webhooks/m365-agent-cli' && url.pathname !== '/webhooks/clippy') {
        sendResponse(res, 404, 'Not Found');
        return;
      }

      const validationToken = url.searchParams.get('validationToken');
      if (validationToken) {
        console.log(`[${new Date().toISOString()}] Received validation token request. Replaying token...`);
        sendResponse(res, 200, validationToken, 'text/plain');
        return;
      }

      if (req.method === 'POST') {
        let body: unknown;
        try {
          body = await readJsonBody(req);
        } catch {
          sendResponse(res, 400, 'Bad Request');
          return;
        }

        const expectedClientState = process.env.GRAPH_CLIENT_STATE;
        const notifications = Array.isArray((body as { value?: unknown })?.value)
          ? (body as { value: unknown[] }).value
          : null;
        if (expectedClientState) {
          const allClientStatesValid =
            !!notifications &&
            notifications.length > 0 &&
            notifications.every(
              (n: unknown) => n && (n as { clientState?: string }).clientState === expectedClientState
            );
          if (!allClientStatesValid) {
            console.warn(
              `[${new Date().toISOString()}] Received Graph notification with invalid or missing clientState.`
            );
            sendResponse(res, 401, 'Invalid clientState');
            return;
          }
        }

        console.log(`[${new Date().toISOString()}] Received Graph notification:`);
        console.log(JSON.stringify(body, null, 2));
        sendResponse(res, 202, 'Accepted');
        return;
      }

      sendResponse(res, 404, 'Not Found');
    } catch (err) {
      console.error('Webhook handler error:', err);
      sendResponse(res, 500, 'Internal Server Error');
    }
  });

  server.listen(port);
}
