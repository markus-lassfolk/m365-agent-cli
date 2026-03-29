import { isIP } from 'node:net';
/**
 * Validates a URL is safe for use as an API endpoint.
 * Blocks SSRF vectors: non-HTTPS protocols, localhost, link-local, and internal IPs.
 */
export function validateUrl(urlString: string, name: string): string {
  let url: URL;
  try {
    url = new URL(urlString);
  } catch {
    throw new Error(`Invalid URL for ${name}: "${urlString}"`);
  }

  if (url.protocol !== 'https:') {
    throw new Error(`${name} must use HTTPS, got: "${urlString}"`);
  }

  let hostname = url.hostname.toLowerCase();

  // Strip brackets from IPv6 addresses before validation
  // WHATWG URL returns IPv6 addresses with brackets (e.g., [::1]), but net.isIP expects bare IPs
  hostname = hostname.replace(/^\[|\]$/g, '');

  // Block localhost variants
  if (hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '::1') {
    throw new Error(`${name} must not point to localhost: "${urlString}"`);
  }

  // Block bare IPv4 addresses — reject all IP literals to prevent internal network access
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  if (isIP(hostname)) {
    throw new Error(`${name} must not be an IP address (use hostname): "${urlString}"`);
  }

  // Block cloud metadata endpoints (common SSRF target)
  if (hostname === 'metadata.google.internal' || hostname.startsWith('169.254.')) {
    throw new Error(`${name} must not point to link-local/metadata endpoint: "${urlString}"`);
  }

  return urlString;
}
