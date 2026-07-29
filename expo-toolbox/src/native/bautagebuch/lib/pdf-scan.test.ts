import assert from 'node:assert/strict';
import { PDFDocument, StandardFonts } from 'pdf-lib';

import { detectedFieldsNeedRescan } from './scan-meta';
import { scanTemplatePdfLite } from './pdf-scan-lite';
import { scanTemplatePdfFull } from './pdf-scan-full';
import { resultHasRects } from './pdf-scan-types';

let passed = 0;

function test(name: string, fn: () => void | Promise<void>) {
  return Promise.resolve(fn()).then(() => {
    passed += 1;
    console.log(`ok - ${name}`);
  });
}

async function buildTextFieldPdf(options: { pages?: number; fieldName?: string } = {}) {
  const pdfDoc = await PDFDocument.create();
  const pages = options.pages ?? 1;
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

  for (let index = 0; index < pages; index += 1) {
    const page = pdfDoc.addPage([595, 842]);
    page.drawText(`Seite ${index + 1}`, { x: 50, y: 800, size: 12, font });
    if (index === pages - 1) {
      const form = pdfDoc.getForm();
      const field = form.createTextField(options.fieldName || 'Text1');
      field.setText('Probe');
      field.addToPage(page, { x: 120, y: 450, width: 130, height: 30 });
    }
  }

  return new Uint8Array(await pdfDoc.save());
}

async function buildCheckboxPdf() {
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([595, 842]);
  const form = pdfDoc.getForm();
  const checkbox = form.createCheckBox('Check1');
  checkbox.addToPage(page, { x: 100, y: 600, width: 18, height: 18 });
  return new Uint8Array(await pdfDoc.save());
}

async function buildPlainPdf() {
  const pdfDoc = await PDFDocument.create();
  pdfDoc.addPage([595, 842]);
  return new Uint8Array(await pdfDoc.save());
}

async function run() {
  await test('scanTemplatePdfLite detects text fields', async () => {
    const bytes = await buildTextFieldPdf();
    const result = await scanTemplatePdfLite(bytes);
    assert.equal(result.fields.length, 1);
    assert.equal(result.fields[0].fieldName, 'Text1');
    assert.equal(result.fields[0].type, 'text');
    assert.ok(result.fields[0].fieldId.includes('text1'));
    assert.equal(result.fields[0].page, 1);
    assert.ok(Array.isArray(result.fields[0].rect));
    assert.equal(result.fields[0].rect!.length, 4);
    assert.equal(result.hasRects, true);
  });

  await test('scanTemplatePdfFull provides page and rect when widgets exist', async () => {
    const bytes = await buildTextFieldPdf({ fieldName: 'Baustelle' });
    const result = await scanTemplatePdfFull(bytes);
    assert.equal(result.fields.length, 1);
    assert.equal(result.fields[0].page, 1);
    assert.ok(Array.isArray(result.fields[0].rect));
    assert.equal(result.fields[0].rect!.length, 4);
    assert.equal(result.hasRects, true);
  });

  await test('scanTemplatePdfLite detects checkbox fields', async () => {
    const bytes = await buildCheckboxPdf();
    const result = await scanTemplatePdfLite(bytes);
    assert.equal(result.fields.length, 1);
    assert.equal(result.fields[0].type, 'checkbox');
  });

  await test('plain PDF without AcroForm throws', async () => {
    const bytes = await buildPlainPdf();
    await assert.rejects(() => scanTemplatePdfLite(bytes), /AcroForm/);
  });

  await test('multi-page PDF lite scan assigns widget page', async () => {
    const bytes = await buildTextFieldPdf({ pages: 2, fieldName: 'FeldP2' });
    const result = await scanTemplatePdfLite(bytes);
    assert.equal(result.pageCount, 2);
    assert.equal(result.fields[0].page, 2);
    assert.ok(Array.isArray(result.fields[0].rect));
  });

  await test('multi-page full scan assigns widget page', async () => {
    const bytes = await buildTextFieldPdf({ pages: 2, fieldName: 'FeldP2' });
    const result = await scanTemplatePdfFull(bytes);
    assert.equal(result.pageCount, 2);
    assert.equal(result.fields[0].page, 2);
  });

  await test('detectedFieldsNeedRescan when fieldId missing', () => {
    assert.equal(
      detectedFieldsNeedRescan([{ fieldId: '', type: 'text', page: 1, rect: [1, 2, 3, 4] }]),
      true
    );
  });

  await test('detectedFieldsNeedRescan when rect missing', () => {
    assert.equal(
      detectedFieldsNeedRescan([
        {
          fieldId: 'a-p1-o1',
          type: 'text',
          page: 1,
          rect: null
        }
      ]),
      true
    );
  });

  await test('detectedFieldsNeedRescan skips manual fields without geometry', () => {
    assert.equal(
      detectedFieldsNeedRescan([
        {
          fieldId: 'manual_1',
          type: 'signature',
          page: 1,
          rect: null,
          source: 'manual'
        }
      ]),
      false
    );
  });

  await test('detectedFieldsNeedRescan when field list empty', () => {
    assert.equal(detectedFieldsNeedRescan([]), false);
  });

  await test('detectedFieldsNeedRescan when complete', () => {
    assert.equal(
      detectedFieldsNeedRescan([
        {
          fieldId: 'a-p1-o1',
          type: 'text',
          page: 1,
          rect: [10, 20, 30, 40]
        }
      ]),
      false
    );
  });

  await test('resultHasRects reflects field rects', () => {
    assert.equal(
      resultHasRects([
        {
          fieldId: 'x',
          fieldName: 'x',
          labelCandidate: 'x',
          type: 'text',
          options: [],
          page: 1,
          orderIndex: 1,
          rect: null
        }
      ]),
      false
    );
    assert.equal(
      resultHasRects([
        {
          fieldId: 'x',
          fieldName: 'x',
          labelCandidate: 'x',
          type: 'text',
          options: [],
          page: 1,
          orderIndex: 1,
          rect: [1, 2, 3, 4]
        }
      ]),
      true
    );
  });

  console.log(`\n${passed} tests passed`);
}

void run();
