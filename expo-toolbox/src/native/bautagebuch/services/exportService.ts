import * as FileSystem from 'expo-file-system/legacy';
import * as Sharing from 'expo-sharing';

import { base64ToUint8Array, bytesToArrayBuffer, uint8ToBase64 } from '../../../lib/binary';
import { buildFinalPdfBytes } from '../lib/pdf-export.js';
import { getRun, getTemplate } from '../db/database';
import { getActiveTemplateBundle } from './templateService';

export async function exportRunPdf(runId: string): Promise<string> {
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

  const filled = await buildFinalPdfBytes({
    templateBlob: {
      arrayBuffer: async () => bytesToArrayBuffer(pdfBytes)
    },
    setupModel,
    runValues: run.values
  });

  const outDir = `${FileSystem.documentDirectory}bautagebuch/exports/`;
  await FileSystem.makeDirectoryAsync(outDir, { intermediates: true });
  const safeTitle = run.title.replace(/[^\w\-äöüÄÖÜß]+/g, '_').slice(0, 80);
  const outPath = `${outDir}${safeTitle || run.runId}.pdf`;
  await FileSystem.writeAsStringAsync(outPath, uint8ToBase64(filled), {
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
