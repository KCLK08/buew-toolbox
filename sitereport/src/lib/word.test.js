import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { buildDocxFilename, exportToDocxData } from './word.js';

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
    columns: [
      { name: 'Bilder', type: 'text', isPhoto: true },
      { name: 'Kilometer', type: 'number', isPhoto: false },
      { name: 'Beschreibung', type: 'text', isPhoto: false },
      { name: 'Status', type: 'text', isPhoto: false }
    ],
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
});
