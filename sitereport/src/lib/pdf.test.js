import { test } from 'node:test';
import assert from 'node:assert/strict';
import { pdfEntryBadgeText, pdfEntryNeedsPhotoArea } from './pdf-entry.js';

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
