/**
 * Build Microsoft Graph chatMessage JSON fragments with @mentions (HTML body).
 * @see https://learn.microsoft.com/en-us/graph/api/channel-post-messages
 */

export interface TeamsAtMention {
  userId: string;
  displayName: string;
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Parse repeatable `--at userId:displayName` (first colon separates id from name). */
export function parseAtSpecs(specs: string[]): TeamsAtMention[] {
  const out: TeamsAtMention[] = [];
  for (const raw of specs) {
    const s = raw.trim();
    const idx = s.indexOf(':');
    if (idx <= 0) {
      throw new Error(`Invalid --at "${raw}": expected userObjectId:displayName`);
    }
    const userId = s.slice(0, idx).trim();
    const displayName = s.slice(idx + 1).trim();
    if (!userId || !displayName) {
      throw new Error(`Invalid --at "${raw}": user id and display name are required`);
    }
    out.push({ userId, displayName });
  }
  return out;
}

/**
 * Turn plain text with literal `@DisplayName` tokens into HTML `<at id="n">…</at>` tags
 * and the Graph `mentions` array. Each mention’s `displayName` must appear in the text as `@displayName` (first occurrence replaced in list order).
 */
export function buildTeamsHtmlBodyWithMentions(
  plainText: string,
  mentions: TeamsAtMention[]
): { body: { contentType: 'html'; content: string }; mentions: unknown[] } {
  let escaped = escapeHtml(plainText);
  const graphMentions: unknown[] = [];

  for (let i = 0; i < mentions.length; i++) {
    const m = mentions[i];
    const plainNeedle = `@${m.displayName}`;
    const needle = `@${escapeHtml(m.displayName)}`;
    if (!escaped.includes(needle)) {
      throw new Error(
        `Text must contain "${plainNeedle}" for mention ${i} (user ${m.userId}). Use @-prefix matching the display name after each --at.`
      );
    }
    // Use a function replacer so `$`-sequences (e.g. `$&`, `$$`) in the display name are
    // inserted literally rather than interpreted by String.prototype.replace.
    escaped = escaped.replace(needle, () => `<at id="${i}">${escapeHtml(m.displayName)}</at>`);
  }

  for (let i = 0; i < mentions.length; i++) {
    const m = mentions[i];
    graphMentions.push({
      id: i,
      mentionText: m.displayName,
      mentioned: {
        user: {
          id: m.userId,
          displayName: m.displayName,
          userIdentityType: 'aadUser'
        }
      }
    });
  }

  return {
    body: { contentType: 'html', content: `<p>${escaped}</p>` },
    mentions: graphMentions
  };
}
