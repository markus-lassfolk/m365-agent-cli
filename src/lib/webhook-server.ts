import { timingSafeEqual } from 'node:crypto';
import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http';

/** Graph notification batches are tiny; cap the body to reject memory-exhaustion payloads. */
const MAX_BODY_BYTES = 256 * 1024;

async function readJsonBody(req: IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = [];
  let total = 0;
  for await (const chunk of req) {
    const buf = chunk as Buffer;
    total += buf.byteLength;
    if (total > MAX_BODY_BYTES) {
      req.destroy();
      throw new Error('Request body too large');
    }
    chunks.push(buf);
  }
  const raw = Buffer.concat(chunks).toString('utf8');
  if (!raw.trim()) return undefined;
  return JSON.parse(raw) as unknown;
}

/** Length-checked constant-time string compare (avoids leaking the secret via timing). */
function safeEqual(a: string, b: string): boolean {
  const ab = Buffer.from(a, 'utf8');
  const bb = Buffer.from(b, 'utf8');
  if (ab.byteLength !== bb.byteLength) return false;
  return timingSafeEqual(ab, bb);
}

/** Redact clientState (a shared secret) from a notification body before logging. */
function redactClientState(body: unknown): unknown {
  if (!body || typeof body !== 'object') return body;
  const value = (body as { value?: unknown }).value;
  if (!Array.isArray(value)) return body;
  return {
    ...(body as Record<string, unknown>),
    value: value.map((n) =>
      n && typeof n === 'object' && 'clientState' in (n as object)
        ? { ...(n as Record<string, unknown>), clientState: '[redacted]' }
        : n
    )
  };
}

function sendResponse(res: ServerResponse, status: number, body: string, contentType = 'text/plain'): void {
  res.writeHead(status, { 'Content-Type': contentType });
  res.end(body);
}

/**
 * Minimal webhook receiver for Graph subscription notifications.
 * Uses Node `http` so the CLI can load under Node/tsx without the Bun runtime.
 *
 * @param port listening port
 * @param host interface to bind (defaults to `WEBHOOK_HOST` env, else all interfaces so
 *   Graph can reach it through a tunnel/reverse proxy)
 */
export function startWebhookServer(port: number = 3000, host: string = process.env.WEBHOOK_HOST || '0.0.0.0'): Server {
  const displayHost = host === '0.0.0.0' || host === '::' ? 'all interfaces' : host;
  console.log(`Starting webhook receiver on ${displayHost}:${port}/webhooks/m365-agent-cli`);
  if (!process.env.GRAPH_CLIENT_STATE) {
    console.warn(
      'Warning: GRAPH_CLIENT_STATE is not set — incoming notifications will NOT be verified. ' +
        'Set GRAPH_CLIENT_STATE (matching your subscription clientState) to reject spoofed notifications.'
    );
  }

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
            notifications.every((n: unknown) => {
              const cs = (n as { clientState?: string } | null)?.clientState;
              return typeof cs === 'string' && safeEqual(cs, expectedClientState);
            });
          if (!allClientStatesValid) {
            console.warn(
              `[${new Date().toISOString()}] Received Graph notification with invalid or missing clientState.`
            );
            sendResponse(res, 401, 'Invalid clientState');
            return;
          }
        }

        console.log(`[${new Date().toISOString()}] Received Graph notification:`);
        console.log(JSON.stringify(redactClientState(body), null, 2));
        sendResponse(res, 202, 'Accepted');
        return;
      }

      sendResponse(res, 404, 'Not Found');
    } catch (err) {
      console.error('Webhook handler error:', err);
      sendResponse(res, 500, 'Internal Server Error');
    }
  });

  // Bound timeouts guard against slow-loris connections holding handlers open.
  server.requestTimeout = 30_000;
  server.headersTimeout = 15_000;

  server.on('error', (err: NodeJS.ErrnoException) => {
    if (err.code === 'EADDRINUSE') {
      console.error(`Error: port ${port} is already in use. Choose another port with --port.`);
    } else if (err.code === 'EACCES') {
      console.error(`Error: permission denied binding to ${displayHost}:${port}. Try a port above 1024.`);
    } else {
      console.error(`Webhook server error: ${err.message}`);
    }
    process.exit(1);
  });

  server.listen(port, host);
  return server;
}
