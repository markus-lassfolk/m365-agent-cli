/**
 * Validates a URL is safe for use as an API endpoint.
 * Blocks SSRF vectors: non-HTTPS protocols, localhost, link-local, and internal IPs.
 */
function validateUrl(urlString: string, name: string): string {
  let url: URL;
  try {
    url = new URL(urlString);
  } catch {
    throw new Error(`Invalid URL for ${name}: "${urlString}"`);
  }

  if (url.protocol !== 'https:') {
    throw new Error(`${name} must use HTTPS, got: "${urlString}"`);
  }

  const hostname = url.hostname.toLowerCase();

  // Block localhost variants
  if (hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '::1' || hostname === '[::1]') {
    throw new Error(`${name} must not point to localhost: "${urlString}"`);
  }

  // Block bare IPv4 addresses (private/internal ranges are a subset of this)
  // Use net.isIP to detect any IP literal — reject all IP addresses to prevent internal network access
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const { isIP } = require('node:net');
  if (isIP(hostname)) {
    throw new Error(`${name} must not be an IP address (use hostname): "${urlString}"`);
  }

  // Block cloud metadata endpoints (common SSRF target)
  if (hostname === 'metadata.google.internal' || hostname.startsWith('169.254.')) {
    throw new Error(`${name} must not point to link-local/metadata endpoint: "${urlString}"`);
  }

  return urlString;
}

export const GRAPH_BASE_URL = validateUrl(
  process.env.GRAPH_BASE_URL || 'https://graph.microsoft.com/v1.0',
  'GRAPH_BASE_URL'
);
