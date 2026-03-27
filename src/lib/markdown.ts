function escapeHtml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function sanitizeLinkUrl(url: string): string {
  const trimmed = url.trim();
  const withoutControlChars = Array.from(trimmed)
    .filter((char) => {
      const code = char.charCodeAt(0);
      return !(
        code <= 0x20 ||
        (code >= 0x7f && code <= 0x9f) ||
        code === 0x200b ||
        code === 0x200c ||
        code === 0x200d ||
        code === 0xfeff
      );
    })
    .join('');
  const decoded = withoutControlChars.replace(/&(#x?[\da-f]+|[a-z]+);?/gi, (entity) => {
    if (/^&#x/i.test(entity)) {
      const value = Number.parseInt(entity.slice(3).replace(/;$/, ''), 16);
      return Number.isFinite(value) ? String.fromCharCode(value) : entity;
    }
    if (/^&#/i.test(entity)) {
      const value = Number.parseInt(entity.slice(2).replace(/;$/, ''), 10);
      return Number.isFinite(value) ? String.fromCharCode(value) : entity;
    }
    const named: Record<string, string> = {
      amp: '&',
      lt: '<',
      gt: '>',
      quot: '"',
      apos: "'"
    };
    const key = entity.slice(1).replace(/;$/, '').toLowerCase();
    return named[key] ?? entity;
  });
  const lower = decoded.toLowerCase();

  if (
    lower.startsWith('javascript:') ||
    lower.startsWith('data:') ||
    lower.startsWith('vbscript:') ||
    lower.startsWith('file:')
  ) {
    return '#';
  }

  return withoutControlChars;
}

/**
 * Convert basic markdown to HTML for email.
 * Supports: bold, italic, links, unordered lists, ordered lists, line breaks.
 */
export function markdownToHtml(text: string): string {
  // Extract and process links first to avoid double-encoding URLs
  const links: Array<{ placeholder: string; html: string }> = [];
  let linkIndex = 0;

  let html = text.replace(/\[([^\]]+)\]\(([^)]+)\)/g, (_, label, rawUrl) => {
    const safeUrl = sanitizeLinkUrl(rawUrl);
    const escapedLabel = escapeHtml(label);
    const escapedUrl = escapeHtml(safeUrl);
    const linkHtml = `<a href="${escapedUrl}">${escapedLabel}</a>`;
    const placeholder = `__LINK_${linkIndex}__`;
    links.push({ placeholder, html: linkHtml });
    linkIndex++;
    return placeholder;
  });

  // Escape HTML in the rest of the text
  html = escapeHtml(html);

  // Restore links
  for (const { placeholder, html: linkHtml } of links) {
    html = html.replace(placeholder, linkHtml);
  }

  // Bold: **text** or __text__
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/__(.+?)__/g, '<strong>$1</strong>');

  // Italic: *text* or _text_ (but not inside words)
  html = html.replace(/(?<!\w)\*([^*]+?)\*(?!\w)/g, '<em>$1</em>');
  html = html.replace(/(?<!\w)_([^_]+?)_(?!\w)/g, '<em>$1</em>');

  // Process lists - need to handle line by line
  const lines = html.split('\n');
  const result: string[] = [];
  let inUnorderedList = false;
  let inOrderedList = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const unorderedMatch = line.match(/^[\s]*[-*]\s+(.+)$/);
    const orderedMatch = line.match(/^[\s]*\d+\.\s+(.+)$/);

    if (unorderedMatch) {
      if (inOrderedList) {
        result.push('</ol>');
        inOrderedList = false;
      }
      if (!inUnorderedList) {
        result.push('<ul>');
        inUnorderedList = true;
      }
      result.push(`<li>${unorderedMatch[1]}</li>`);
    } else if (orderedMatch) {
      if (inUnorderedList) {
        result.push('</ul>');
        inUnorderedList = false;
      }
      if (!inOrderedList) {
        result.push('<ol>');
        inOrderedList = true;
      }
      result.push(`<li>${orderedMatch[1]}</li>`);
    } else {
      // Close any open lists
      if (inUnorderedList) {
        result.push('</ul>');
        inUnorderedList = false;
      }
      if (inOrderedList) {
        result.push('</ol>');
        inOrderedList = false;
      }
      result.push(line);
    }
  }

  // Close any remaining open lists
  if (inUnorderedList) {
    result.push('</ul>');
  }
  if (inOrderedList) {
    result.push('</ol>');
  }

  html = result.join('\n');

  // Convert line breaks to <br> (but not inside lists)
  // Split by list tags, process non-list parts
  html = html.replace(/\n(?!<\/?[uo]l>|<\/?li>)/g, '<br>\n');

  // Wrap in basic HTML structure for email
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; line-height: 1.5; }
  a { color: #0066cc; }
  ul, ol { margin: 10px 0; padding-left: 20px; }
  li { margin: 5px 0; }
</style>
</head>
<body>
${html}
</body>
</html>`;
}

/**
 * Check if text contains markdown formatting.
 */
export function hasMarkdown(text: string): boolean {
  // Check for common markdown patterns
  return /\*\*.+?\*\*|__.+?__|(?<!\w)\*[^*]+?\*(?!\w)|(?<!\w)_[^_]+?_(?!\w)|\[.+?\]\(.+?\)|^[\s]*[-*]\s+|^[\s]*\d+\.\s+/m.test(
    text
  );
}
