import { describe, expect, it } from 'bun:test';
import { MailTemplateError, parseTemplateVars, renderMailTemplate } from './mail-template.js';

describe('parseTemplateVars', () => {
  it('parses name=value pairs', () => {
    expect(parseTemplateVars(['name=Alice', 'company=Acme'])).toEqual({ name: 'Alice', company: 'Acme' });
  });

  it('splits on the first = only, so values may contain =', () => {
    expect(parseTemplateVars(['formula=a=b=c'])).toEqual({ formula: 'a=b=c' });
  });

  it('allows an empty value', () => {
    expect(parseTemplateVars(['name='])).toEqual({ name: '' });
  });

  it('throws for a pair missing =', () => {
    expect(() => parseTemplateVars(['bad'])).toThrow(MailTemplateError);
    expect(() => parseTemplateVars(['bad'])).toThrow(/Invalid --var/);
  });

  it('throws for an empty name', () => {
    expect(() => parseTemplateVars(['=value'])).toThrow(/empty name/);
  });
});

describe('renderMailTemplate', () => {
  it('substitutes a simple placeholder', () => {
    expect(renderMailTemplate('Hello {{name}}!', { name: 'Alice' })).toBe('Hello Alice!');
  });

  it('substitutes multiple distinct placeholders', () => {
    expect(
      renderMailTemplate('{{greeting}} {{name}}, from {{company}}.', { greeting: 'Hi', name: 'Bob', company: 'Acme' })
    ).toBe('Hi Bob, from Acme.');
  });

  it('uses the placeholder default when no var is supplied', () => {
    expect(renderMailTemplate('Hello {{name|there}}!', {})).toBe('Hello there!');
  });

  it('prefers a supplied var over the default', () => {
    expect(renderMailTemplate('Hello {{name|there}}!', { name: 'Alice' })).toBe('Hello Alice!');
  });

  it('allows an empty default (renders as empty string)', () => {
    expect(renderMailTemplate('Sig:{{sig|}}', {})).toBe('Sig:');
  });

  it('leaves unrelated text and repeated placeholders alone', () => {
    expect(renderMailTemplate('{{name}} and {{name}} again', { name: 'X' })).toBe('X and X again');
  });

  it('throws MailTemplateError listing every unresolved placeholder, once each', () => {
    expect(() => renderMailTemplate('{{a}} {{b}} {{a}}', {})).toThrow(MailTemplateError);
    try {
      renderMailTemplate('{{a}} {{b}} {{a}}', {});
      throw new Error('expected renderMailTemplate to throw');
    } catch (err) {
      expect(err).toBeInstanceOf(MailTemplateError);
      expect((err as Error).message).toContain('a, b');
    }
  });

  it('ignores extra vars not referenced by the template', () => {
    expect(renderMailTemplate('Hello {{name}}!', { name: 'Alice', unused: 'x' })).toBe('Hello Alice!');
  });

  it('throws on a malformed placeholder (leading digit) instead of passing it through literally', () => {
    expect(() => renderMailTemplate('Hello {{1name}}!', {})).toThrow(MailTemplateError);
    expect(() => renderMailTemplate('Hello {{1name}}!', {})).toThrow(/malformed placeholder/);
  });

  it('throws on a malformed placeholder (hyphenated name)', () => {
    expect(() => renderMailTemplate('Hello {{my-name}}!', { 'my-name': 'x' })).toThrow(/malformed placeholder/);
  });

  it('throws on an empty placeholder', () => {
    expect(() => renderMailTemplate('Hello {{}}!', {})).toThrow(/malformed placeholder/);
  });

  it('does not flag a valid placeholder as malformed once resolved', () => {
    expect(renderMailTemplate('Hello {{name}}!', { name: 'Alice' })).toBe('Hello Alice!');
  });

  it('does not flag a {{...}}-shaped substring inside a substituted --var value as malformed (bug regression)', () => {
    expect(renderMailTemplate('Welcome, {{name}}!', { name: 'Team {{Placeholder}}' })).toBe(
      'Welcome, Team {{Placeholder}}!'
    );
  });
});
