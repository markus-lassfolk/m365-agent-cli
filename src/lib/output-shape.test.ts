import { describe, expect, it } from 'bun:test';
import { formatNdjson, parseFieldsOption, projectFields, shapeRows } from './output-shape.js';

describe('parseFieldsOption', () => {
  it('splits and trims comma-separated fields', () => {
    expect(parseFieldsOption('subject, from.emailAddress.address ,id')).toEqual([
      'subject',
      'from.emailAddress.address',
      'id'
    ]);
  });

  it('drops empty segments from trailing/double commas', () => {
    expect(parseFieldsOption('subject,,id,')).toEqual(['subject', 'id']);
  });

  it('returns undefined for unset or empty input', () => {
    expect(parseFieldsOption(undefined)).toBeUndefined();
    expect(parseFieldsOption('')).toBeUndefined();
    expect(parseFieldsOption('   ')).toBeUndefined();
    expect(parseFieldsOption(',,,')).toBeUndefined();
  });
});

describe('projectFields', () => {
  it('picks top-level fields', () => {
    expect(projectFields({ id: '1', subject: 'hi', extra: 'x' }, ['id', 'subject'])).toEqual({
      id: '1',
      subject: 'hi'
    });
  });

  it('picks nested dot-path fields, preserving nesting', () => {
    const row = { from: { emailAddress: { address: 'a@b.com', name: 'A' } }, subject: 'hi' };
    expect(projectFields(row, ['from.emailAddress.address', 'subject'])).toEqual({
      from: { emailAddress: { address: 'a@b.com' } },
      subject: 'hi'
    });
  });

  it('silently omits paths that do not exist', () => {
    expect(projectFields({ id: '1' }, ['id', 'missing.path'])).toEqual({ id: '1' });
  });

  it('returns non-object values (null, primitives, arrays) unchanged', () => {
    expect(projectFields(null, ['id'])).toBeNull();
    expect(projectFields('x', ['id'])).toBe('x');
    expect(projectFields([1, 2], ['id'])).toEqual([1, 2]);
  });

  it('treats an empty fields list as "keep nothing"', () => {
    expect(projectFields({ id: '1', subject: 'hi' }, [])).toEqual({});
  });
});

describe('shapeRows', () => {
  const rows = [
    { id: '1', subject: 'a', from: { emailAddress: { address: 'a@x.com' } } },
    { id: '2', subject: 'b', from: { emailAddress: { address: 'b@x.com' } } }
  ];

  it('returns rows unchanged when fields is undefined or empty', () => {
    expect(shapeRows(rows, undefined)).toBe(rows);
    expect(shapeRows(rows, [])).toBe(rows);
  });

  it('projects every row when fields is set', () => {
    expect(shapeRows(rows, ['id', 'from.emailAddress.address'])).toEqual([
      { id: '1', from: { emailAddress: { address: 'a@x.com' } } },
      { id: '2', from: { emailAddress: { address: 'b@x.com' } } }
    ]);
  });
});

describe('formatNdjson', () => {
  it('prints one compact JSON object per line', () => {
    const out = formatNdjson([{ id: '1' }, { id: '2' }]);
    expect(out).toBe('{"id":"1"}\n{"id":"2"}');
    expect(out.split('\n')).toHaveLength(2);
  });

  it('returns an empty string for an empty list', () => {
    expect(formatNdjson([])).toBe('');
  });
});
