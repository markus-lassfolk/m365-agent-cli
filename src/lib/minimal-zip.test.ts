import { afterEach, describe, expect, test } from 'bun:test';
import { execFileSync } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { crc32, createStoredZip } from './minimal-zip.js';

describe('crc32', () => {
  test('matches the well-known "The quick brown fox..." CRC-32 value', () => {
    const buf = Buffer.from('The quick brown fox jumps over the lazy dog', 'utf8');
    expect(crc32(buf).toString(16)).toBe('414fa339');
  });

  test('CRC of empty buffer is 0', () => {
    expect(crc32(Buffer.alloc(0))).toBe(0);
  });
});

describe('createStoredZip', () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) await rm(tmpDir, { recursive: true, force: true }).catch(() => {});
  });

  test('produces a zip with correct local/central/EOCD signatures', () => {
    const zip = createStoredZip([{ name: 'diagnostic.json', content: Buffer.from('{"ok":true}', 'utf8') }]);
    expect(zip.readUInt32LE(0)).toBe(0x04034b50); // local file header
    expect(zip.indexOf(Buffer.from('diagnostic.json'))).toBeGreaterThan(0);
    expect(zip.indexOf(Buffer.from('{"ok":true}'))).toBeGreaterThan(0);
    expect(zip.readUInt32LE(zip.length - 22)).toBe(0x06054b50); // EOCD is the last 22 bytes
  });

  test('a real unzip tool can list and extract the archive contents byte-for-byte', async () => {
    tmpDir = await mkdtemp(join(tmpdir(), 'm365-zip-test-'));
    const content = Buffer.from(JSON.stringify({ hello: 'world', n: 42 }, null, 2), 'utf8');
    const zip = createStoredZip([{ name: 'diagnostic.json', content }]);
    const zipPath = join(tmpDir, 'bundle.zip');
    await writeFile(zipPath, zip);

    const listing = execFileSync('unzip', ['-l', zipPath], { encoding: 'utf8' });
    expect(listing).toContain('diagnostic.json');

    execFileSync('unzip', ['-o', zipPath, '-d', tmpDir], { encoding: 'utf8' });
    const extracted = await Bun.file(join(tmpDir, 'diagnostic.json')).text();
    expect(extracted).toBe(content.toString('utf8'));
  });

  test("Python's zipfile module (independent implementation) parses the archive without CRC errors", async () => {
    tmpDir = await mkdtemp(join(tmpdir(), 'm365-zip-py-test-'));
    const content = Buffer.from('a'.repeat(5000), 'utf8'); // exercise more than one CRC table pass
    const zip = createStoredZip([{ name: 'diagnostic.json', content }]);
    const zipPath = join(tmpDir, 'bundle.zip');
    await writeFile(zipPath, zip);

    const script = `
import zipfile, sys
with zipfile.ZipFile(sys.argv[1]) as zf:
    bad = zf.testzip()
    assert bad is None, f"corrupt member: {bad}"
    data = zf.read('diagnostic.json')
    assert len(data) == ${content.length}, len(data)
print("OK")
`;
    const out = execFileSync('python3', ['-c', script, zipPath], { encoding: 'utf8' });
    expect(out.trim()).toBe('OK');
  });

  test('multiple entries all round-trip correctly', async () => {
    tmpDir = await mkdtemp(join(tmpdir(), 'm365-zip-multi-'));
    const zip = createStoredZip([
      { name: 'a.json', content: Buffer.from('{"a":1}') },
      { name: 'b.txt', content: Buffer.from('hello') }
    ]);
    const zipPath = join(tmpDir, 'bundle.zip');
    await writeFile(zipPath, zip);
    const listing = execFileSync('unzip', ['-l', zipPath], { encoding: 'utf8' });
    expect(listing).toContain('a.json');
    expect(listing).toContain('b.txt');
  });
});
