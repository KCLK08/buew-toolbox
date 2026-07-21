import * as FileSystem from 'expo-file-system/legacy';
import * as Sharing from 'expo-sharing';

import { base64ToUint8Array, bytesToArrayBuffer, uint8ToBase64 } from '../../../lib/binary';
import { buildFinalPdfBytes } from '../lib/pdf-export.js';
import { buildPhotoDocPdfBytes, mergeBtbWithPhotoDoc } from '../lib/photo-doc.js';
import { getRun, getTemplate } from '../db/database';
import { getActiveTemplateBundle } from './templateService';
import { readPhotoBytes } from './photoDocService';

export type BautagebuchExportMode = 'btb' | 'photo' | 'merged';

export async function exportRunPdf(runId: string, mode: BautagebuchExportMode = 'merged'): Promise<string> {
  const run = await getRun(runId);
  if (!run) {
    throw new Error('BTB-Lauf nicht gefunden.');
  }

  const { setupModel } = await getActiveTemplateBundle();
  const templateRecord = await getTemplate(run.templateId);
  if (!templateRecord?.pdfPath) {
    throw new Error('PDF-Vorlage fehlt.');
  }

  const base64 = await FileSystem.readAsStringAsync(templateRecord.pdfPath, {
    encoding: FileSystem.EncodingType.Base64
  });
  const pdfBytes = base64ToUint8Array(base64);

  let outputBytes: Uint8Array;
  const safeTitle = run.title.replace(/[^\w\-äöüÄÖÜß]+/g, '_').slice(0, 80);

  if (mode === 'photo') {
    const photoEntries = await buildPhotoEntries(run);
    outputBytes = await (buildPhotoDocPdfBytes as unknown as (input: {
      title: string;
      entries: Array<{ photoBlob: { bytes: Uint8Array; mimeType: string } }>;
    }) => Promise<Uint8Array>)({
      title: run.title,
      entries: photoEntries
    });
  } else {
    const filled = await buildFinalPdfBytes({
      templateBlob: {
        arrayBuffer: async () => bytesToArrayBuffer(pdfBytes)
      },
      setupModel,
      runValues: run.values
    });

    if (mode === 'merged') {
      const photoEntries = await buildPhotoEntries(run);
      const photoEnabled = run.photoDoc?.enabled === true || photoEntries.length > 0;
      const merged = await (mergeBtbWithPhotoDoc as unknown as (input: {
        btbPdfBytes: Uint8Array;
        photoDocEnabled: boolean;
        photoEntries: Array<{ photoBlob: { bytes: Uint8Array; mimeType: string } }>;
      }) => Promise<{ bytes: Uint8Array }>)({
        btbPdfBytes: filled,
        photoDocEnabled: photoEnabled,
        photoEntries
      });
      outputBytes = merged.bytes;
    } else {
      outputBytes = filled;
    }
  }

  const suffix = mode === 'photo' ? '_Fotodoku' : mode === 'merged' ? '_komplett' : '';
  const outDir = `${FileSystem.documentDirectory}bautagebuch/exports/`;
  await FileSystem.makeDirectoryAsync(outDir, { intermediates: true });
  const outPath = `${outDir}${safeTitle || run.runId}${suffix}.pdf`;
  await FileSystem.writeAsStringAsync(outPath, uint8ToBase64(outputBytes), {
    encoding: FileSystem.EncodingType.Base64
  });

  if (await Sharing.isAvailableAsync()) {
    await Sharing.shareAsync(outPath, {
      mimeType: 'application/pdf',
      dialogTitle: run.title
    });
  }

  return outPath;
}

async function buildPhotoEntries(run: Awaited<ReturnType<typeof getRun>>) {
  const entries = run?.photoDoc?.entries || [];
  const prepared = [];
  for (const entry of entries) {
    if (!entry.localPath) continue;
    const bytes = await readPhotoBytes(entry.localPath);
    prepared.push({
      photoBlob: { bytes, mimeType: entry.mimeType || 'image/jpeg' }
    });
  }
  return prepared;
}
