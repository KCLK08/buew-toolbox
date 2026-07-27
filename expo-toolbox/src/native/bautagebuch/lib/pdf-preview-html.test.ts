import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import {
  buildFieldPreviewHtml,
  buildPdfJsInlineScript,
  buildSimplePdfPreviewHtml,
  escapeInlineScript,
  PDF_PREVIEW_LOAD_ERROR,
  PDFJS_VERSION
} from './pdf-preview-html';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const assetsDir = path.resolve(testDir, '../../../../assets/pdfjs');

const mockAssets = {
  pdfJsSource: 'window.pdfjsLib = { GlobalWorkerOptions: {}, getDocument: () => ({ promise: Promise.resolve({}) }) };',
  workerSrc: 'file:///offline/pdf.worker.min.js'
};

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

function assertNoCdnUrls(html: string) {
  assert.doesNotMatch(html, /cdnjs\.cloudflare\.com/i, 'CDN URL must not appear in preview HTML');
  assert.doesNotMatch(html, /ajax\/libs\/pdf\.js/i, 'pdf.js CDN path must not appear in preview HTML');
}

test('escapeInlineScript neutralizes closing script tags', () => {
  const escaped = escapeInlineScript('foo</script>bar');
  assert.equal(escaped, 'foo<\\/script>bar');
});

test('buildPdfJsInlineScript inlines local core instead of CDN script src', () => {
  const snippet = buildPdfJsInlineScript('/* local pdfjs core */');
  assert.match(snippet, /<script>\s*\/\* local pdfjs core \*\/\s*<\/script>/);
  assert.doesNotMatch(snippet, /<script src=/);
});

test('buildFieldPreviewHtml uses bundled assets and offline worker URL', () => {
  const html = buildFieldPreviewHtml({
    base64: 'UEZERg==',
    ...mockAssets,
    mode: 'mapping',
    highlights: [{ fieldId: 'f1', page: 1, rect: [10, 20, 30, 40] }]
  });

  assertNoCdnUrls(html);
  assert.match(html, /window\.pdfjsLib/);
  assert.match(html, /GlobalWorkerOptions\.workerSrc = 'file:\/\/\/offline\/pdf\.worker\.min\.js'/);
  assert.match(html, /PDF preview boot failed/);
  assert.match(html, /PDF preview render failed/);
  assert.match(html, new RegExp(PDF_PREVIEW_LOAD_ERROR.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
});

test('buildSimplePdfPreviewHtml rejects legacy base64-only signature', () => {
  assert.throws(
    () => buildSimplePdfPreviewHtml('UEZERg==' as unknown as Parameters<typeof buildSimplePdfPreviewHtml>[0]),
    /requires bundled pdf\.js assets/
  );
});

test('buildSimplePdfPreviewHtml embeds local worker and renderer boot helpers', () => {
  const html = buildSimplePdfPreviewHtml({
    base64: 'UEZERg==',
    ...mockAssets
  });

  assertNoCdnUrls(html);
  assert.match(html, /GlobalWorkerOptions\.workerSrc = 'file:\/\/\/offline\/pdf\.worker\.min\.js'/);
  assert.match(html, /showPreviewError\('PDF preview boot failed'/);
});

test('offline simulation: bundled asset files exist for pdf.js 3.11.174', () => {
  const corePath = path.join(assetsDir, 'pdf.min.bundle');
  const workerPath = path.join(assetsDir, 'pdf.worker.min.bundle');
  const core = readFileSync(corePath, 'utf8');
  const worker = readFileSync(workerPath, 'utf8');

  assert.ok(core.length > 100_000, 'pdf.min.bundle should be bundled locally');
  assert.ok(worker.length > 500_000, 'pdf.worker.min.bundle should be bundled locally');
  assert.match(core, /pdfjsLib|pdf\.js/i);
  assert.match(worker, /worker/i);
  assert.equal(PDFJS_VERSION, '3.11.174');
});

test('offline simulation: local core can be injected without network script tags', () => {
  const core = readFileSync(path.join(assetsDir, 'pdf.min.bundle'), 'utf8');
  const html = buildFieldPreviewHtml({
    base64: 'UEZERg==',
    pdfJsSource: core.slice(0, 4096),
    workerSrc: 'file:///var/mobile/Containers/Data/Application/offline/pdf.worker.min.js'
  });

  assertNoCdnUrls(html);
  assert.doesNotMatch(html, /<script src=/);
  assert.match(html, /GlobalWorkerOptions\.workerSrc = 'file:\/\/\/var\/mobile/);
});

console.log(`\n${passed} tests passed`);
