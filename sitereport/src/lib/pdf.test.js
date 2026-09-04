import { test } from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument } from 'pdf-lib';
import {
  estimatePdfHeaderRemaining,
  layoutPdfEntryFlow,
  pdfEntryBadgeText,
  pdfEntryNeedsPhotoArea,
  pdfMaxHeaderHeight,
  pdfPageBodyHeight,
  pdfTwoUpCardBudget,
  planPdfEntryPlacement
} from './pdf-entry.js';
import { exportToPdfData } from './pdf.js';

test('pdf badges use Eintrag numbering instead of Bild', () => {
  assert.equal(pdfEntryBadgeText(0), 'Eintrag 1');
  assert.equal(pdfEntryBadgeText(1), 'Eintrag 2');
  assert.equal(pdfEntryBadgeText(9), 'Eintrag 10');
});

test('entries without photos skip the PDF image area', () => {
  assert.equal(pdfEntryNeedsPhotoArea({}), false);
  assert.equal(pdfEntryNeedsPhotoArea({ fields: { Beschreibung: 'Nur Text' } }), false);
  assert.equal(pdfEntryNeedsPhotoArea({ photoBlobs: [] }), false);
  assert.equal(pdfEntryNeedsPhotoArea({ photoBlob: null, photoBlobs: [] }), false);
});

test('entries with photos still get a PDF image area', () => {
  assert.equal(pdfEntryNeedsPhotoArea({ photoBlob: 'blob' }), true);
  assert.equal(pdfEntryNeedsPhotoArea({ photoBlobs: ['a', 'b'] }), true);
});

test('first PDF entry always stays on the header page', () => {
  const twoUp = pdfTwoUpCardBudget();
  const plan = planPdfEntryPlacement({
    remaining: 420,
    isFirstDocumentPage: true,
    entriesOnPage: 0,
    photoCount: 3,
    naturalCardHeight: 520,
    chromeHeight: 120
  });
  assert.equal(plan.stayOnPage, true);
  assert.equal(plan.forceFit, true);
  assert.ok(plan.maxCardHeight <= twoUp + 0.01);
  assert.ok(plan.maxCardHeight <= 420);
});

test('first PDF page force-fits a second 1-photo entry when leftover is enough', () => {
  const plan = planPdfEntryPlacement({
    remaining: 180,
    isFirstDocumentPage: true,
    entriesOnPage: 1,
    photoCount: 1,
    naturalCardHeight: 300,
    chromeHeight: 110
  });
  assert.equal(plan.stayOnPage, true);
  assert.equal(plan.forceFit, true);
});

test('first PDF page skips a second entry when leftover is too small', () => {
  const plan = planPdfEntryPlacement({
    remaining: 80,
    isFirstDocumentPage: true,
    entriesOnPage: 1,
    photoCount: 1,
    naturalCardHeight: 300,
    chromeHeight: 110
  });
  assert.equal(plan.stayOnPage, false);
});

test('continuation pages size a normal entry for two-up', () => {
  const twoUp = pdfTwoUpCardBudget();
  const plan = planPdfEntryPlacement({
    remaining: 770,
    isFirstDocumentPage: false,
    entriesOnPage: 0,
    photoCount: 1,
    naturalCardHeight: 500,
    chromeHeight: 110
  });
  assert.equal(plan.stayOnPage, true);
  assert.ok(plan.maxCardHeight <= twoUp + 0.01);
});

test('a second continuation entry stays when it fits', () => {
  const plan = planPdfEntryPlacement({
    remaining: 390,
    isFirstDocumentPage: false,
    entriesOnPage: 1,
    photoCount: 1,
    naturalCardHeight: 300,
    chromeHeight: 110
  });
  assert.equal(plan.stayOnPage, true);
});

test('many-photo entries can take a whole continuation page', () => {
  const plan = planPdfEntryPlacement({
    remaining: 770,
    isFirstDocumentPage: false,
    entriesOnPage: 0,
    photoCount: 6,
    naturalCardHeight: 620,
    chromeHeight: 110
  });
  assert.equal(plan.stayOnPage, true);
  assert.ok(plan.maxCardHeight > pdfTwoUpCardBudget());
});

test('PDF packing keeps two short entries on page one, then two per later page', () => {
  const columns = [
    { name: 'Kilometer' },
    { name: 'Beschreibung' },
    { name: 'Status' }
  ];
  const entries = Array.from({ length: 5 }, (_, i) => ({
    fields: { Kilometer: String(i), Beschreibung: `Text ${i}`, Status: 'offen' }
  }));
  const flow = layoutPdfEntryFlow({
    entries,
    tableColumns: columns,
    photoSizesForEntry: () => [],
    headerRemaining: estimatePdfHeaderRemaining({
      protocolTitle: 'Protokoll',
      protocolDescription: 'Kurz',
      attendees: 'Max'
    })
  });
  assert.equal(flow.length, 5);
  assert.equal(flow[0].pageBreakBefore, false);
  assert.equal(flow[1].pageBreakBefore, false);
  assert.equal(flow[2].pageBreakBefore, true);
  assert.equal(flow[3].pageBreakBefore, false);
  assert.equal(flow[4].pageBreakBefore, true);
});

test('a long header still leaves room for the first entry', () => {
  const remaining = estimatePdfHeaderRemaining({
    protocolTitle: 'Sehr langes Protokoll',
    protocolDescription: Array.from({ length: 20 }, (_, i) => `Beschreibung Zeile ${i + 1}`).join('\n'),
    attendees: Array.from({ length: 12 }, (_, i) => `Person ${i + 1}`).join('\n')
  });
  assert.ok(remaining >= pdfTwoUpCardBudget() - 0.01);
  assert.ok(pdfMaxHeaderHeight() < pdfPageBodyHeight());
});

test('one- and two-photo entries are packed two per continuation page', () => {
  const columns = [
    { name: 'Kilometer' },
    { name: 'Beschreibung' },
    { name: 'Status' }
  ];
  const entries = Array.from({ length: 4 }, (_, i) => ({
    fields: { Kilometer: String(i), Beschreibung: `Text ${i}`, Status: 'offen' }
  }));
  const photos = (count) => Array.from({ length: count }, () => ({ width: 1600, height: 900 }));
  for (const photoCount of [1, 2]) {
    const flow = layoutPdfEntryFlow({
      entries,
      tableColumns: columns,
      photoSizesForEntry: () => photos(photoCount),
      headerRemaining: estimatePdfHeaderRemaining({
        protocolTitle: 'Protokoll',
        protocolDescription: 'Kurz',
        attendees: 'Max'
      })
    });
    assert.equal(flow.length, 4, `${photoCount} photos`);
    assert.equal(flow[0].pageBreakBefore, false, `${photoCount} photos first stays`);
    assert.equal(flow[1].pageBreakBefore, false, `${photoCount} photos second stays`);
    assert.equal(flow[2].pageBreakBefore, true, `${photoCount} photos page 2`);
    assert.equal(flow[3].pageBreakBefore, false, `${photoCount} photos two-up`);
    const twoUp = pdfTwoUpCardBudget();
    assert.ok(flow[0].maxCardHeight <= twoUp + 1);
    assert.ok(flow[2].maxCardHeight <= twoUp + 1);
  }
});

const PDF_COLUMNS = [
  { name: 'Bilder', isPhoto: true },
  { name: 'Kilometer', isPhoto: false },
  { name: 'Beschreibung', isPhoto: false },
  { name: 'Status', isPhoto: false }
];

test('PDF keeps a single entry on the header page even with a long header', async () => {
  const result = await exportToPdfData({
    protocolTitle: 'Langprotokoll',
    projectName: 'Nord',
    protocolDate: '04-09-2026',
    protocolDescription: Array.from({ length: 16 }, (_, i) => `Beschreibung Zeile ${i + 1}`).join('\n'),
    attendees: Array.from({ length: 10 }, (_, i) => `Person ${i + 1}`).join('\n'),
    logoDataUrl: '',
    columns: PDF_COLUMNS,
    entries: [
      { id: 'e1', fields: { Kilometer: '1', Beschreibung: 'Erster Eintrag', Status: 'offen' } }
    ]
  });
  const pdf = await PDFDocument.load(Buffer.from(result.base64, 'base64'));
  assert.equal(pdf.getPageCount(), 1);
  assert.equal(result.stats.exportedEntries, 1);
});

test('PDF packs four short entries onto two pages', async () => {
  const result = await exportToPdfData({
    protocolTitle: 'Protokoll',
    projectName: 'Nord',
    protocolDate: '04-09-2026',
    protocolDescription: 'Kurz',
    attendees: 'Max',
    logoDataUrl: '',
    columns: PDF_COLUMNS,
    entries: Array.from({ length: 4 }, (_, i) => ({
      id: `e${i + 1}`,
      fields: { Kilometer: String(i + 1), Beschreibung: `Text ${i + 1}`, Status: 'offen' }
    }))
  });
  const pdf = await PDFDocument.load(Buffer.from(result.base64, 'base64'));
  assert.equal(pdf.getPageCount(), 2);
  assert.equal(result.stats.exportedEntries, 4);
});
