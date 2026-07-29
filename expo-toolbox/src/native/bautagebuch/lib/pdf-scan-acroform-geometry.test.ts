import assert from 'node:assert/strict';
import { PDFDocument, StandardFonts } from 'pdf-lib';

import {
  countScanFieldGeometry,
  extractFieldWidgetMetadata,
  logAcroFormImportStats
} from './pdf-scan-acroform-geometry';
import { scanTemplatePdfLite } from './pdf-scan-lite';

let passed = 0;

function test(name: string, fn: () => void | Promise<void>) {
  return Promise.resolve(fn()).then(() => {
    passed += 1;
    console.log(`ok - ${name}`);
  });
}

async function buildTwoPageFieldPdf() {
  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const page1 = pdfDoc.addPage([595, 842]);
  page1.drawText('Seite 1', { x: 50, y: 800, size: 12, font });
  const page2 = pdfDoc.addPage([595, 842]);
  page2.drawText('Seite 2', { x: 50, y: 800, size: 12, font });
  const form = pdfDoc.getForm();
  const field = form.createTextField('PageTwoField');
  field.setText('Wert');
  field.addToPage(page2, { x: 80, y: 420, width: 160, height: 28 });
  return { pdfDoc, bytes: new Uint8Array(await pdfDoc.save()) };
}

async function run() {
  await test('extractFieldWidgetMetadata resolves page and rect from widgets', async () => {
    const { pdfDoc, bytes } = await buildTwoPageFieldPdf();
    const [field] = pdfDoc.getForm().getFields();
    const widget = extractFieldWidgetMetadata(pdfDoc, field, 0);
    assert.equal(widget.page, 2);
    assert.ok(Array.isArray(widget.rect));
    assert.equal(widget.rect!.length, 4);
    assert.ok(widget.rect![2] > widget.rect![0]);
    assert.ok(widget.rect![3] > widget.rect![1]);

    const scan = await scanTemplatePdfLite(bytes);
    assert.equal(scan.fields[0].page, 2);
    assert.deepEqual(scan.fields[0].rect, widget.rect);
  });

  await test('countScanFieldGeometry separates fields with and without rects', () => {
    const stats = countScanFieldGeometry([
      { rect: [1, 2, 3, 4] },
      { rect: null },
      { rect: [10, 20, 30, 40] }
    ]);
    assert.equal(stats.total, 3);
    assert.equal(stats.withGeometry, 2);
    assert.equal(stats.withoutGeometry, 1);
  });

  await test('logAcroFormImportStats prints summary without throwing', () => {
    const original = console.log;
    const lines: string[] = [];
    console.log = (...args: unknown[]) => {
      lines.push(args.map(String).join(' '));
    };
    try {
      logAcroFormImportStats([{ rect: [1, 2, 3, 4] }, { rect: null }]);
      assert.ok(lines.some((line) => line.includes('AcroForm fields: 2')));
      assert.ok(lines.some((line) => line.includes('Fields with geometry: 1')));
      assert.ok(lines.some((line) => line.includes('Fields without geometry: 1')));
    } finally {
      console.log = original;
    }
  });

  console.log(`\n${passed} pdf-scan-acroform-geometry tests passed`);
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
