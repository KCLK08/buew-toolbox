import assert from 'node:assert/strict';

import {
  prepareMultilinePdfText,
  rectSizeFromPdfBox,
  shouldAutoFitPdfField,
  wrapTextToLines
} from './pdf-text-fit.js';

let passed = 0;

function test(name, fn) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

const C4_RECT = { width: 248, height: 96 };
const REAL_C4_RECT = { width: 236.4, height: 267.5 };

test('shouldAutoFitPdfField matches Leistungsblock C4/C5', () => {
  assert.equal(shouldAutoFitPdfField({ tableId: 'table_detail_blocks', columnId: 'c4' }), true);
  assert.equal(shouldAutoFitPdfField({ tableId: 'table_detail_blocks', columnId: 'c5' }), true);
  assert.equal(shouldAutoFitPdfField({ fieldName: 'Text66' }), true);
  assert.equal(shouldAutoFitPdfField({ tableId: 'table_detail_blocks', columnId: 'c1' }), false);
  assert.equal(shouldAutoFitPdfField({ tableId: 'table_main_personal', columnId: 'c4' }), false);
});

test('rectSizeFromPdfBox reads widget and pdf.js boxes', () => {
  assert.deepEqual(rectSizeFromPdfBox({ width: 100, height: 40 }), { width: 100, height: 40 });
  assert.deepEqual(rectSizeFromPdfBox([10, 20, 110, 60]), { width: 100, height: 40 });
});

test('short C4 text keeps a readable start size', () => {
  const prepared = prepareMultilinePdfText({
    text: 'Gleisbau Abschnitt A',
    rect: C4_RECT,
    startFontSize: 12
  });
  assert.equal(prepared.fontSize, 12);
  assert.match(prepared.text, /Gleisbau Abschnitt A/);
  assert.equal(prepared.text.includes('…'), false);
});

test('medium C4 text stays fully visible and may shrink slightly', () => {
  const text =
    'a) Ausgeführte Arbeiten: Schienenausbau, Schwellenwechsel und Stopfarbeiten im Abschnitt Nord.\nb) Geräte: Zweiwegebagger, Stopfmaschine, Schienensäge.';
  const prepared = prepareMultilinePdfText({
    text,
    rect: C4_RECT,
    startFontSize: 12
  });
  assert.ok(prepared.fontSize <= 12);
  assert.ok(prepared.fontSize >= 6);
  assert.equal(prepared.text.includes('Stopfarbeiten'), true);
  assert.equal(prepared.text.includes('Schienensäge'), true);
  assert.equal(prepared.text.includes('…'), false);
});

test('very long C4 text shrinks font and keeps every word', () => {
  const sentences = Array.from({ length: 8 }, (_, index) => {
    return `Abschnitt ${index + 1}: Gleisrichtung, Schotterung und Vermessung der ausgeführten Arbeiten.`;
  });
  const text = sentences.join(' ');
  const prepared = prepareMultilinePdfText({
    text,
    rect: C4_RECT,
    startFontSize: 12
  });

  assert.ok(prepared.fontSize < 12);
  assert.ok(prepared.fontSize >= 6);
  assert.equal(prepared.text.includes('…'), false);
  for (const sentence of sentences) {
    for (const word of sentence.split(' ')) {
      assert.equal(prepared.text.includes(word), true, `missing word: ${word}`);
    }
  }

  const boxHeight = C4_RECT.height - Math.min(6, Math.max(2, C4_RECT.height * 0.14)) * 2;
  assert.ok(prepared.lines.length * prepared.lineHeight <= boxHeight + 0.01);
});

test('explicit line breaks stay as paragraph boundaries', () => {
  const prepared = prepareMultilinePdfText({
    text: 'Erste Zeile bleibt oben.\nZweite Zeile nach Umbruch.\nDritte Zeile mit noch mehr Inhalt zur Kontrolle.',
    rect: C4_RECT,
    startFontSize: 12
  });
  const joined = prepared.lines.join('\n');
  assert.match(joined, /Erste Zeile/);
  assert.match(joined, /Zweite Zeile/);
  assert.match(joined, /Dritte Zeile/);
});

test('wrapTextToLines does not overflow the measured column width', () => {
  const maxWidth = 180;
  const fontSize = 10;
  const lines = wrapTextToLines('Wort '.repeat(40).trim(), maxWidth, fontSize);
  const maxChars = Math.max(4, Math.floor(maxWidth / Math.max(1, fontSize * 0.52)));
  assert.ok(lines.length > 1);
  for (const line of lines) {
    assert.ok(line.length <= maxChars, `"${line}" exceeds ${maxChars} chars`);
  }
});

test('realistic tall C4 column keeps a readable size for moderate text', () => {
  const text =
    'a) Ausgeführte Arbeiten: Schienenausbau, Schwellenwechsel und Stopfarbeiten im Abschnitt Nord.\nb) Geräte: Zweiwegebagger, Stopfmaschine, Schienensäge.';
  const prepared = prepareMultilinePdfText({
    text,
    rect: REAL_C4_RECT,
    startFontSize: 12
  });
  assert.equal(prepared.fontSize, 12);
  assert.equal(prepared.text.includes('Stopfarbeiten'), true);
  assert.equal(prepared.text.includes('…'), false);
});

test('realistic tall C4 column shrinks only when the full text would overflow', () => {
  const text = Array.from(
    { length: 16 },
    (_, index) =>
      `Abschnitt ${index + 1}: Gleisrichtung, Schotterung, Vermessung und Dokumentation der ausgeführten Arbeiten inklusive verwendeter Maschinen.`
  ).join(' ');
  const prepared = prepareMultilinePdfText({
    text,
    rect: REAL_C4_RECT,
    startFontSize: 12
  });
  assert.ok(prepared.fontSize < 12);
  assert.ok(prepared.fontSize >= 6);
  assert.equal(prepared.text.includes('…'), false);
  for (const word of ['Gleisrichtung', 'Schotterung', 'Vermessung', 'Maschinen']) {
    assert.equal(prepared.text.includes(word), true);
  }
  const boxHeight = REAL_C4_RECT.height - Math.min(6, Math.max(2, REAL_C4_RECT.height * 0.14)) * 2;
  assert.ok(prepared.lines.length * prepared.lineHeight <= boxHeight + 0.01);
});

test('several long C4 entries remain complete after fit', () => {
  const entries = [
    'Gleis 3: Schwellenauswechslung auf 120 m, inkl. Schotterung und Durcharbeitung.',
    'Weiche 12: Zungenvorrichtung geprüft, Schmierung erneuert, Protokoll erstellt.',
    'Baustellenlogistik: Materialanfuhr, Zwischenlagerung und Abtransport Altstoffe.'
  ];
  const prepared = prepareMultilinePdfText({
    text: entries.join('\n\n'),
    rect: C4_RECT,
    startFontSize: 12
  });
  for (const entry of entries) {
    for (const word of entry.split(' ')) {
      assert.equal(prepared.text.includes(word), true, `missing word: ${word}`);
    }
  }
  assert.equal(prepared.text.includes('…'), false);
});

console.log(`\n${passed} tests passed`);
