import { describe, expect, it } from 'bun:test';
import { hasMarkdown, markdownToHtml } from './markdown.js';

describe('markdownToHtml', () => {
  it('sanitizes javascript urls in markdown links', () => {
    const html = markdownToHtml('[click](javascript:alert(1))');
    expect(html).toContain('<a href="#">click</a>');
    expect(html).not.toContain('javascript:');
  });

  it('sanitizes obfuscated javascript urls in markdown links', () => {
    const html = markdownToHtml('[click](java\nscript:alert(1))');
    expect(html).toContain('<a href="#">click</a>');
    expect(html).not.toContain('javascript:');
  });

  it('escapes html in link labels and urls', () => {
    const html = markdownToHtml('[<b>hi</b>](https://example.com?q=%3Cscript%3E)');
    expect(html).toContain('<a href="https://example.com?q=%3Cscript%3E">&lt;b&gt;hi&lt;/b&gt;</a>');
  });

  it('sanitizes data, vbscript, and file URLs', () => {
    expect(markdownToHtml('[d](data:text/html,xx)')).toContain('href="#"');
    expect(markdownToHtml('[v](vbscript:msgbox(1))')).toContain('href="#"');
    expect(markdownToHtml('[f](file:///etc/passwd)')).toContain('href="#"');
  });

  it('decodes numeric and named entities in link href before scheme check', () => {
    const html = markdownToHtml('[x](&#106;avascript:alert(1))');
    expect(html).toContain('href="#"');
  });

  it('does not corrupt link URLs or labels containing emphasis characters', () => {
    // Emphasis passes must not rewrite `_`/`*` inside a restored link.
    const html = markdownToHtml('[report](https://host.com/_a_b_)');
    expect(html).toContain('href="https://host.com/_a_b_"');
    expect(html).not.toContain('<em>');
    const html2 = markdownToHtml('see [my_file](https://x.com/p) and *emph*');
    expect(html2).toContain('href="https://x.com/p"');
    expect(html2).toContain('my_file');
    // Emphasis outside links still renders.
    expect(html2).toContain('<em>emph</em>');
  });

  it('renders bold, italic, lists, line breaks, and wraps document', () => {
    const md = '**B** and *I*\n\n- a\n- b\n\n1. x\n2. y\n\nplain';
    const html = markdownToHtml(md);
    expect(html).toContain('<strong>B</strong>');
    expect(html).toContain('<em>I</em>');
    expect(html).toContain('<ul>');
    expect(html).toContain('<ol>');
    expect(html).toContain('<!DOCTYPE html>');
    expect(html).toContain('<br>');
  });
});

describe('hasMarkdown', () => {
  it('detects common patterns', () => {
    expect(hasMarkdown('**x**')).toBe(true);
    expect(hasMarkdown('_y_')).toBe(true);
    expect(hasMarkdown('[a](u)')).toBe(true);
    expect(hasMarkdown('- item')).toBe(true);
    expect(hasMarkdown('1. item')).toBe(true);
    expect(hasMarkdown('plain')).toBe(false);
  });
});
