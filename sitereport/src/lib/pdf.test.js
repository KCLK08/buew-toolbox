import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  estimatePdfHeaderRemaining,
  layoutPdfEntryFlow,
  pdfEntryBadgeText,
  pdfEntryNeedsPhotoArea,
  pdfTwoUpCardBudget,
  planPdfEntryPlacement
} from './pdf-entry.js';

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
  assert.equal(plan.maxCardHeight, 420);
});

test('first PDF page does not require a second entry', () => {
  const plan = planPdfEntryPlacement({
    remaining: 180,
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
