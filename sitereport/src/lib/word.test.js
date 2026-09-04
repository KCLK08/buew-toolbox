import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { layoutPdfEntryFlow, estimatePdfHeaderRemaining, pdfTwoUpCardBudget } from './pdf-entry.js';
import { buildDocxFilename, exportToDocxData } from './word.js';

const TEXT_COLUMNS = [
  { name: 'Bilder', type: 'text', isPhoto: true },
  { name: 'Kilometer', type: 'number', isPhoto: false },
  { name: 'Beschreibung', type: 'text', isPhoto: false },
  { name: 'Status', type: 'text', isPhoto: false }
];

function textEntries(count) {
  return Array.from({ length: count }, (_, i) => ({
    id: `e${i + 1}`,
    fields: {
      Kilometer: String(i + 1),
      Beschreibung: `Text ohne Foto ${i + 1}`,
      Status: 'offen'
    }
  }));
}

function expectedWordPageBreaks(entries, meta = {}) {
  const tableColumns = TEXT_COLUMNS.filter((col) => !col.isPhoto);
  const flow = layoutPdfEntryFlow({
    entries,
    tableColumns,
    photoSizesForEntry: () => [],
    headerRemaining: Math.max(pdfTwoUpCardBudget(), estimatePdfHeaderRemaining(meta) - 36)
  });
  return flow.filter((item) => item.pageBreakBefore).length;
}

test('buildDocxFilename uses project, date and .docx', () => {
  assert.equal(buildDocxFilename('Nord Bau', '03-09-2026'), 'Nord Bau_03-09-2026.docx');
  assert.ok(buildDocxFilename('', '').endsWith('.docx'));
});

test('Word export includes Eintrag badges and field values without a photo placeholder', async () => {
  const result = await exportToDocxData({
    protocolTitle: 'Testprotokoll Word',
    projectName: 'Projekt Nord',
    protocolDate: '04-09-2026',
    protocolDescription: 'Beschreibung Zeile 1',
    attendees: 'Max Mustermann',
    logoDataUrl: '',
    columns: TEXT_COLUMNS,
    entries: [
      {
        id: 'e1',
        fields: { Kilometer: '12', Beschreibung: 'Text ohne Foto', Status: 'offen' }
      },
      {
        id: 'e2',
        fields: { Kilometer: '3', Beschreibung: 'Zweiter Eintrag', Status: 'erledigt' }
      }
    ]
  });

  assert.equal(result.filename, 'Projekt Nord_04-09-2026.docx');
  assert.equal(result.stats.format, 'docx');
  assert.equal(result.stats.exportedEntries, 2);
  assert.ok(result.base64.length > 100);

  const zip = await JSZip.loadAsync(Buffer.from(result.base64, 'base64'));
  const xml = await zip.file('word/document.xml').async('string');
  assert.match(xml, /Eintrag 1/);
  assert.match(xml, /Eintrag 2/);
  assert.match(xml, /Text ohne Foto/);
  assert.match(xml, /Zweiter Eintrag/);
  assert.match(xml, /Beschreibung Zeile 1/);
  assert.doesNotMatch(xml, /Kein Bild vorhanden/);
  assert.doesNotMatch(xml, /Bild 1/);
  assert.match(xml, /w:fill="F7F8F9"/);
  assert.match(xml, /<w:tbl[\s\S]*<w:tbl[\s\S]*Eintrag 1/);
});

test('Word uses the same page-break packing as the PDF', async () => {
  const entries = textEntries(5);
  const meta = {
    protocolTitle: 'Testprotokoll Word',
    protocolDescription: 'Kurz',
    attendees: 'Max'
  };
  const result = await exportToDocxData({
    ...meta,
    projectName: 'Projekt Nord',
    protocolDate: '04-09-2026',
    logoDataUrl: '',
    columns: TEXT_COLUMNS,
    entries
  });

  const zip = await JSZip.loadAsync(Buffer.from(result.base64, 'base64'));
  const xml = await zip.file('word/document.xml').async('string');
  const pageBreaks = (xml.match(/w:type="page"/g) || []).length;
  const expected = expectedWordPageBreaks(entries, meta);
  assert.equal(expected, 2);
  assert.equal(pageBreaks, expected);
  assert.match(xml, /Eintrag 5/);
});

test('two short Word entries stay on the first page like the PDF', async () => {
  const entries = textEntries(2);
  const result = await exportToDocxData({
    protocolTitle: 'Kurzprotokoll',
    projectName: 'Projekt Nord',
    protocolDate: '04-09-2026',
    protocolDescription: 'Kurz',
    attendees: 'Max',
    logoDataUrl: '',
    columns: TEXT_COLUMNS,
    entries
  });
  const zip = await JSZip.loadAsync(Buffer.from(result.base64, 'base64'));
  const xml = await zip.file('word/document.xml').async('string');
  assert.equal((xml.match(/w:type="page"/g) || []).length, 0);
});

test('Word keeps the first entry with a long header instead of page-breaking', async () => {
  const result = await exportToDocxData({
    protocolTitle: 'Langprotokoll',
    projectName: 'Projekt Nord',
    protocolDate: '04-09-2026',
    protocolDescription: Array.from({ length: 16 }, (_, i) => `Beschreibung Zeile ${i + 1}`).join('\n'),
    attendees: Array.from({ length: 10 }, (_, i) => `Person ${i + 1}`).join('\n'),
    logoDataUrl: '',
    columns: TEXT_COLUMNS,
    entries: textEntries(2)
  });
  const zip = await JSZip.loadAsync(Buffer.from(result.base64, 'base64'));
  const xml = await zip.file('word/document.xml').async('string');
  assert.match(xml, /Eintrag 1/);
  assert.match(xml, /w:keepNext/);
  assert.equal((xml.match(/w:type="page"/g) || []).length, 0);
});
