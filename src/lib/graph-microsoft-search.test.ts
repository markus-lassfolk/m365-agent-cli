import { describe, expect, test } from 'bun:test';
import {
  buildMicrosoftSearchRequest,
  deepMergeSearchRequest,
  flattenMicrosoftSearchHits,
  type MicrosoftSearchQueryResponse
} from './graph-microsoft-search.js';

describe('deepMergeSearchRequest', () => {
  test('merges nested query object', () => {
    const base = { query: { queryString: 'a' }, from: 0 };
    const merged = deepMergeSearchRequest(base, { query: { queryTemplate: '{searchTerms}' } });
    expect(merged).toEqual({
      query: { queryString: 'a', queryTemplate: '{searchTerms}' },
      from: 0
    });
  });

  test('overlay replaces arrays', () => {
    const merged = deepMergeSearchRequest({ fields: ['a'] }, { fields: ['b', 'c'] });
    expect(merged.fields).toEqual(['b', 'c']);
  });

  test('ignores prototype-polluting keys from the overlay', () => {
    const overlay = JSON.parse('{"query":{"queryString":"x"},"__proto__":{"polluted":true},"constructor":{"bad":1}}');
    const merged = deepMergeSearchRequest({ from: 0 }, overlay);
    expect(merged).toEqual({ from: 0, query: { queryString: 'x' } });
    // No prototype pollution of plain objects.
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
    expect(Object.hasOwn(merged, '__proto__')).toBe(false);
  });
});

describe('buildMicrosoftSearchRequest', () => {
  test('applies requestPatch', () => {
    const r = buildMicrosoftSearchRequest({
      entityTypes: ['message'],
      queryString: 'x',
      from: 1,
      size: 10,
      requestPatch: { region: 'US' }
    });
    expect(r.entityTypes).toEqual(['message']);
    expect(r.query).toEqual({ queryString: 'x' });
    expect(r.from).toBe(1);
    expect(r.size).toBe(10);
    expect(r.region).toBe('US');
  });
});

describe('flattenMicrosoftSearchHits', () => {
  test('extracts hits from nested value', () => {
    const res: MicrosoftSearchQueryResponse = {
      value: [
        {
          hitsContainers: [
            {
              hits: [
                {
                  rank: 1,
                  hitId: 'h1',
                  summary: 'S',
                  resource: {
                    '@odata.type': '#microsoft.graph.driveItem',
                    id: 'item1',
                    webUrl: 'https://example.com/x',
                    name: 'Doc.docx'
                  }
                }
              ]
            }
          ]
        }
      ]
    };
    const hits = flattenMicrosoftSearchHits(res);
    expect(hits).toHaveLength(1);
    expect(hits[0]).toMatchObject({
      rank: 1,
      hitId: 'h1',
      summary: 'S',
      entityType: 'driveItem',
      id: 'item1',
      webUrl: 'https://example.com/x',
      name: 'Doc.docx'
    });
  });

  test('empty response yields empty array', () => {
    expect(flattenMicrosoftSearchHits({})).toEqual([]);
  });
});
