import { describe, expect, it } from 'bun:test';
import { markdownToHtml } from './markdown.js';

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
});
