import assert from 'node:assert/strict';
import { PDFDocument } from 'pdf-lib';

import { applyPdfFieldValue } from './setup-model.js';
import { prepareMultilinePdfText } from './pdf-text-fit.js';

let passed = 0;

function test(name, fn) {
  return Promise.resolve()
    .then(fn)
    .then(() => {
      passed += 1;
      console.log(`ok - ${name}`);
    });
}

function readFieldFontSize(field) {
  const da = field?.acroField?.getDefaultAppearance?.() || '';
  const match = String(da).match(/\/\S+\s+([\d.]+)\s+Tf/);
  return match ? Number(match[1]) : null;
}

async function createC4Field(width = 248, height = 96) {
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([595, 842]);
  const form = pdfDoc.getForm();
  const field = form.createTextField('Text66');
  field.addToPage(page, { x: 40, y: 420, width, height });
  field.acroField.setDefaultAppearance('/Helv 12 Tf 0 g');
  return { pdfDoc, field, form };
}

await test('C4 export keeps short text at 12pt', async () => {
  const { field } = await createC4Field();
  applyPdfFieldValue(field, 'text', 'Gleisbau Abschnitt A', {
    fieldName: 'Text66',
    tableId: 'table_detail_blocks',
    columnId: 'c4'
  });
  assert.equal(field.getText(), 'Gleisbau Abschnitt A');
  assert.equal(readFieldFontSize(field), 12);
});

await test('C4 export shrinks font and keeps the full long text', async () => {
  const text = Array.from(
    { length: 8 },
    (_, index) =>
      `Abschnitt ${index + 1}: Gleisrichtung, Schotterung und Vermessung der ausgeführten Arbeiten.`
  ).join(' ');
  const { field } = await createC4Field();
  applyPdfFieldValue(field, 'text', text, {
    fieldName: 'Text66',
    tableId: 'table_detail_blocks',
    columnId: 'c4'
  });

  const exported = field.getText() || '';
  const fontSize = readFieldFontSize(field);
  const expected = prepareMultilinePdfText({
    text,
    rect: { width: 248, height: 96 },
    startFontSize: 12
  });

  assert.ok(fontSize < 12);
  assert.equal(fontSize, expected.fontSize);
  assert.equal(exported.includes('…'), false);
  for (const word of text.split(' ')) {
    assert.equal(exported.includes(word), true, `missing export word: ${word}`);
  }
  assert.equal(exported, expected.text);
});

await test('C4 preview and export share the same fitted layout', async () => {
  const text =
    'a) Ausgeführte Arbeiten: Schienenausbau und Schwellenwechsel.\nb) Geräte: Zweiwegebagger und Stopfmaschine.';
  const preview = prepareMultilinePdfText({
    text,
    rect: { width: 248, height: 96 },
    startFontSize: 12
  });
  const { field } = await createC4Field();
  applyPdfFieldValue(field, 'text', text, {
    fieldName: 'Text66',
    tableId: 'table_detail_blocks',
    columnId: 'c4'
  });
  assert.equal(field.getText(), preview.text);
  assert.equal(readFieldFontSize(field), preview.fontSize);
});

await test('tall C4 field keeps 12pt for moderate text and still wraps', async () => {
  const text =
    'a) Ausgeführte Arbeiten: Schienenausbau, Schwellenwechsel und Stopfarbeiten.\nb) Geräte: Zweiwegebagger und Stopfmaschine.';
  const { field } = await createC4Field(236.4, 267.5);
  const preview = prepareMultilinePdfText({
    text,
    rect: { width: 236.4, height: 267.5 },
    startFontSize: 12
  });
  applyPdfFieldValue(field, 'text', text, {
    fieldName: 'Text66',
    tableId: 'table_detail_blocks',
    columnId: 'c4'
  });
  assert.equal(preview.fontSize, 12);
  assert.equal(readFieldFontSize(field), 12);
  assert.equal(field.getText(), preview.text);
  assert.match(field.getText() || '', /Stopfarbeiten/);
});

console.log(`\n${passed} tests passed`);
