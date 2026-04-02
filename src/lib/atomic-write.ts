import { randomBytes } from 'node:crypto';
import { mkdir, rename, unlink, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';

/**
 * Write UTF-8 text atomically (temp file in same directory, then rename).
 * Avoids readers observing partial writes and pairs well with validation before persist.
 */
export async function atomicWriteUtf8File(targetPath: string, data: string, mode: number): Promise<void> {
  const dir = dirname(targetPath);
  await mkdir(dir, { recursive: true, mode: 0o700 });
  const tmp = join(dir, `.${randomBytes(16).toString('hex')}.tmp`);
  try {
    await writeFile(tmp, data, { encoding: 'utf8', mode });
    await rename(tmp, targetPath);
  } catch (err) {
    await unlink(tmp).catch(() => {});
    throw err;
  }
}
