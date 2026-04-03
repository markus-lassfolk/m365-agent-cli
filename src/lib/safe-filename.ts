import { basename } from 'node:path';

const INVALID_CHARS = /[/\\?%*:|"<>]/g;

/**
 * Reduces Graph/API attachment names to a single safe path component under a trusted output directory.
 * Mitigates path traversal and CodeQL "network data written to file" findings.
 */
export function safeAttachmentFileName(raw: string | undefined | null, fallback: string): string {
  const trimmed = String(raw ?? '').trim();
  const base = basename(trimmed.length > 0 ? trimmed : fallback);
  const noTraversal = base.replace(/\.\./g, '_');
  const s = noTraversal.replace(INVALID_CHARS, '_').trim();
  const out = s.length > 0 ? s : fallback;
  return out.length > 255 ? out.slice(0, 255) : out;
}

/** HTTP(S) URL safe for `.url` InternetShortcut content (blocks CRLF / scheme injection). */
export function safeHttpUrlForInternetShortcut(url: string): string | null {
  const u = url.trim();
  if (!u || u.includes('\r') || u.includes('\n') || u.includes('\0')) {
    return null;
  }
  try {
    const parsed = new URL(u);
    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
      return null;
    }
    return parsed.href;
  } catch {
    return null;
  }
}
