import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import {
  buildFieldPreviewHtml,
  buildPdfJsInlineScript,
  buildPdfWorkerInlineScript,
  buildScrollableFieldPreviewHtml,
  buildSimplePdfPreviewHtml,
  escapeInlineScript,
  PDF_PREVIEW_LOAD_ERROR,
  PDFJS_VERSION
} from './pdf-preview-html';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const assetsDir = path.resolve(testDir, '../../../../assets/pdfjs');

const mockAssets = {
  pdfJsSource: 'window.pdfjsLib = { GlobalWorkerOptions: {}, getDocument: () => ({ promise: Promise.resolve({}) }) };',
  workerSrc: 'file:///offline/pdf.worker.min.js',
  workerSource: 'globalThis.pdfjsWorker = { WorkerMessageHandler: function() {} };'
};

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

async function testAsync(name: string, fn: () => Promise<void>) {
  await fn();
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

test('buildPdfWorkerInlineScript inlines worker before core bootstrap', () => {
  const snippet = buildPdfWorkerInlineScript('/* pdf.worker */');
  assert.match(snippet, /<script>\s*\/\* pdf\.worker \*\/\s*<\/script>/);
  assert.doesNotMatch(snippet, /<script src=/);
});

test('buildFieldPreviewHtml inlines worker then core and uses fake-worker boot', () => {
  const html = buildFieldPreviewHtml({
    base64: 'UEZERg==',
    ...mockAssets,
    mode: 'mapping',
    highlights: [{ fieldId: 'f1', page: 1, rect: [10, 20, 30, 40] }]
  });

  assertNoCdnUrls(html);
  const workerPos = html.indexOf('globalThis.pdfjsWorker');
  const corePos = html.indexOf('window.pdfjsLib');
  assert.ok(workerPos >= 0 && corePos > workerPos, 'worker script must appear before core script');
  assert.match(html, /PdfPreviewWorkerDisabled/);
  assert.match(html, /globalThis\.pdfjsWorker\?\.WorkerMessageHandler/);
  assert.match(html, /GlobalWorkerOptions\.workerSrc = 'file:\/\/\/offline\/pdf\.worker\.min\.js'/);
  assert.match(html, /PDF preview boot failed/);
  assert.match(html, /PDF preview render failed/);
  assert.match(html, new RegExp(PDF_PREVIEW_LOAD_ERROR.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
  assert.doesNotMatch(html, /URL\.createObjectURL\(workerBlob\)/);
});

test('buildSimplePdfPreviewHtml rejects legacy base64-only signature', () => {
  assert.throws(
    () => buildSimplePdfPreviewHtml('UEZERg==' as unknown as Parameters<typeof buildSimplePdfPreviewHtml>[0]),
    /requires bundled pdf\.js assets/
  );
});

test('buildSimplePdfPreviewHtml uses scrollable multi-page layout without field overlay', () => {
  const html = buildSimplePdfPreviewHtml({
    base64: 'UEZERg==',
    ...mockAssets
  });

  assertNoCdnUrls(html);
  assert.match(html, /PdfPreviewWorkerDisabled/);
  assert.match(html, /GlobalWorkerOptions\.workerSrc = 'file:\/\/\/offline\/pdf\.worker\.min\.js'/);
  assert.match(html, /showPreviewError\('PDF preview boot failed'/);
  assert.match(html, /renderPageSheet/);
  assert.match(html, /Scrollen · Zwei Finger zum Zoomen/);
  assert.doesNotMatch(html, /id="overlay"/);
  assert.doesNotMatch(html, /setPage/);
});

test('buildScrollableFieldPreviewHtml uses scrollable layout with field overlays', () => {
  const html = buildScrollableFieldPreviewHtml({
    base64: 'UEZERg==',
    ...mockAssets,
    highlights: [{ fieldId: 'f1', fieldName: 'Text1', page: 1, rect: [10, 10, 50, 50] }],
    highlightActive: true
  });

  assert.match(html, /renderPageSheet/);
  assert.match(html, /\.overlay/);
  assert.match(html, /__applyPreviewCommand/);
  assert.match(html, /matchesActive/);
  assert.match(html, /Scrollen · Zwei Finger zum Zoomen/);
  assert.doesNotMatch(html, /setPage/);
  assert.doesNotMatch(html, /Seite/);
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

test('offline simulation: inlined worker HTML disables native Worker bootstrap', () => {
  const core = readFileSync(path.join(assetsDir, 'pdf.min.bundle'), 'utf8');
  const worker = readFileSync(path.join(assetsDir, 'pdf.worker.min.bundle'), 'utf8');
  const html = buildFieldPreviewHtml({
    base64: 'UEZERg==',
    pdfJsSource: core,
    workerSource: worker,
    workerSrc: 'file:///var/mobile/Containers/Data/Application/offline/pdf.worker.min.js'
  });

  assertNoCdnUrls(html);
  assert.doesNotMatch(html, /<script src=/);
  assert.match(html, /PdfPreviewWorkerDisabled/);
  assert.match(html, /pdfjsWorker\?\.WorkerMessageHandler/);
  const headClose = html.indexOf('</head>');
  const workerScriptOpen = html.indexOf('<script>', html.indexOf('<head>'));
  const coreScriptOpen = html.indexOf('<script>', workerScriptOpen + 1);
  assert.ok(workerScriptOpen >= 0 && coreScriptOpen > workerScriptOpen && coreScriptOpen < headClose);
});

void (async () => {
  await testAsync('offline simulation: inlined worker + fake Worker loads PDF', async () => {
    const core = readFileSync(path.join(assetsDir, 'pdf.min.bundle'), 'utf8');
    const worker = readFileSync(path.join(assetsDir, 'pdf.worker.min.bundle'), 'utf8');

    globalThis.window = globalThis as unknown as Window & typeof globalThis;
    eval(worker);
    eval(core);
    globalThis.Worker = class {
      constructor() {
        throw new Error('Workers disabled');
      }
    } as unknown as typeof Worker;

    type PdfJsLib = {
      GlobalWorkerOptions: { workerSrc: string };
      getDocument: (opts: { data: Uint8Array }) => { promise: Promise<{ numPages: number }> };
    };

    const lib = (globalThis as { pdfjsLib?: PdfJsLib }).pdfjsLib;
    assert.ok(lib, 'pdfjsLib should exist after eval');
    assert.ok((globalThis as { pdfjsWorker?: { WorkerMessageHandler?: unknown } }).pdfjsWorker?.WorkerMessageHandler);

    lib!.GlobalWorkerOptions.workerSrc = 'file:///fake/worker.js';
    const emptyPdf = Buffer.from(
      'JVBERi0xLjQKJeLjz9MKMSAwIG9iago8PAovVHlwZSAvQ2F0YWxvZwovUGFnZXMgMiAwIFIKPj4KZW5kb2JqCjIgMCBvYmoKPDwKL1R5cGUgL1BhZ2VzCi9LaWRzIFszIDAgUl0KL0NvdW50IDEKL01lZGlhQm94IFswIDAgNjEyIDc5Ml0KPj4KZW5vb2JqCjMgMCBvYmoKPDwKL1R5cGUgL1BhZ2UKL1BhcmVudCAyIDAgUgovUmVzb3VyY2VzIDw8Cj4+Ci9NZWRpYUJveCBbMCAwIDYxMiA3OTJdCi9Db250ZW50cyA0IDAgUgo+PgplbmRvYmoKNCAwIG9iago8PAovTGVuZ3RoIDQ0Cj4+CnN0cmVhbQpCVAovRjEgMjQgVGYgMTAwIDcwMCBUZCAoSGVsbG8pIFRqIEVUCmVuZHN0cmVhbQplbmRvYmoKeHJlZgowIDUKMDAwMDAwMDAwMCA2NTUzNSBmIAowMDAwMDAwMDA5IDAwMDAwIG4gCjAwMDAwMDAwNTggMDAwMDAgbiAKMDAwMDAwMDExNSAwMDAwMCBuIAowMDAwMDAwMjA2IDAwMDAwIG4gCnRyYWlsZXIKPDwKL1NpemUgNQovUm9vdCAxIDAgUgo+PgpzdGFydHh4ZWYKMjkzCiUlRU9G',
      'base64'
    );

    const doc = await lib!.getDocument({ data: new Uint8Array(emptyPdf) }).promise;
    assert.equal(doc.numPages, 1);
  });

  console.log(`\n${passed} tests passed`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
