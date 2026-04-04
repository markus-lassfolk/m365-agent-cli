import { describe, expect, test } from 'bun:test';
import { buildGraphSendMailPayload } from './graph-send-mail.js';

describe('buildGraphSendMailPayload', () => {
  test('includes referenceAttachment entries when link attachments are set', () => {
    const { message } = buildGraphSendMailPayload({
      to: ['a@b.com'],
      subject: 's',
      body: 'hi',
      html: false,
      referenceAttachments: [{ name: 'Doc', sourceUrl: 'https://example.com/x' }]
    });
    const atts = message.attachments as Record<string, unknown>[];
    expect(atts).toHaveLength(1);
    expect(atts[0]['@odata.type']).toBe('#microsoft.graph.referenceAttachment');
    expect(atts[0].name).toBe('Doc');
    expect(atts[0].sourceUrl).toBe('https://example.com/x');
  });

  test('merges file and reference attachments', () => {
    const { message } = buildGraphSendMailPayload({
      to: ['a@b.com'],
      subject: 's',
      body: 'hi',
      html: false,
      fileAttachments: [{ name: 'f.bin', contentType: 'application/octet-stream', contentBytes: 'YQ==' }],
      referenceAttachments: [{ name: 'Link', sourceUrl: 'https://example.com/' }]
    });
    const atts = message.attachments as Record<string, unknown>[];
    expect(atts).toHaveLength(2);
    expect(atts[0]['@odata.type']).toBe('#microsoft.graph.fileAttachment');
    expect(atts[1]['@odata.type']).toBe('#microsoft.graph.referenceAttachment');
  });
});
