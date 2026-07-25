import * as FileSystem from 'expo-file-system/legacy';
import * as Sharing from 'expo-sharing';

import { base64ToUint8Array, bytesToArrayBuffer, uint8ToBase64 } from '../../../lib/binary';
import { nowIso } from '../../../lib/ids';
import { fotodokuTitleFromBtbTitle } from '../lib/btb-naming';
import { buildFinalPdfBytes } from '../lib/pdf-export.js';
import { buildPhotoDocPdfBytes, mergeBtbWithPhotoDoc } from '../lib/photo-doc.js';
import { deleteExport, getExport, getRun, getSetupModel, getTemplate, upsertExportByRun } from '../db/database';
import { readPhotoBytes } from './photoDocService';

export type BautagebuchExportMode = 'btb' | 'photo' | 'merged';

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

async function buildRunPdfBytes(runId: string, mode: BautagebuchExportMode): Promise<{
  bytes: Uint8Array;
  safeTitle: string;
  suffix: string;
}> {
  const run = await getRun(runId);
  if (!run) {
    throw new Error('BTB-Lauf nicht gefunden.');
  }

  const setupModel = await getSetupModel(run.templateId);
  if (!setupModel) {
    throw new Error('Setup-Modell für diese Vorlage fehlt.');
  }
  const templateRecord = await getTemplate(run.templateId);
  if (!templateRecord?.pdfPath) {
    throw new Error('PDF-Vorlage fehlt.');
  }

  const base64 = await FileSystem.readAsStringAsync(templateRecord.pdfPath, {
    encoding: FileSystem.EncodingType.Base64
  });
  const pdfBytes = base64ToUint8Array(base64);
  const exportTitle = mode === 'photo' ? fotodokuTitleFromBtbTitle(run.title) : run.title;
  const safeTitle = exportTitle.replace(/[^\w\-äöüÄÖÜß]+/g, '_').slice(0, 80);

  if (mode === 'photo') {
    const photoEntries = await buildPhotoEntries(run);
    const bytes = await (buildPhotoDocPdfBytes as unknown as (input: {
      title: string;
      entries: Array<{ photoBlob: { bytes: Uint8Array; mimeType: string } }>;
    }) => Promise<Uint8Array>)({
      title: exportTitle,
      entries: photoEntries
    });
    return { bytes, safeTitle, suffix: '' };
  }

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
    return { bytes: merged.bytes, safeTitle, suffix: '_komplett' };
  }

  return { bytes: filled, safeTitle, suffix: '' };
}

async function writeRunExport(
  runId: string,
  mode: BautagebuchExportMode,
  options: { persistExport?: boolean } = {}
): Promise<string> {
  const { persistExport = true } = options;
  const { bytes, safeTitle, suffix } = await buildRunPdfBytes(runId, mode);
  const outDir = persistExport
    ? `${FileSystem.documentDirectory}bautagebuch/exports/`
    : `${FileSystem.documentDirectory}bautagebuch/previews/`;
  await FileSystem.makeDirectoryAsync(outDir, { intermediates: true });
  const fileName = persistExport ? `${safeTitle || runId}${suffix}.pdf` : `${runId}_preview.pdf`;
  const outPath = `${outDir}${fileName}`;
  await FileSystem.writeAsStringAsync(outPath, uint8ToBase64(bytes), {
    encoding: FileSystem.EncodingType.Base64
  });
  if (persistExport) {
    await upsertExportByRun({
      exportId: `export_${runId}`,
      runId,
      fileName,
      filePath: outPath,
      exportedAt: nowIso()
    });
  }
  return outPath;
}

async function sharePdf(path: string, title: string) {
  if (await Sharing.isAvailableAsync()) {
    await Sharing.shareAsync(path, {
      mimeType: 'application/pdf',
      dialogTitle: title
    });
  }
}

export async function exportRunPdf(runId: string, mode: BautagebuchExportMode = 'merged'): Promise<string> {
  const run = await getRun(runId);
  if (!run) throw new Error('BTB-Lauf nicht gefunden.');
  const outPath = await writeRunExport(runId, mode);
  await sharePdf(outPath, run.title);
  return outPath;
}

export async function previewRunPdf(runId: string): Promise<string> {
  const run = await getRun(runId);
  if (!run) throw new Error('BTB-Lauf nicht gefunden.');
  const outPath = await writeRunExport(runId, 'btb');
  await sharePdf(outPath, `${run.title} (Vorschau)`);
  return outPath;
}

/** Generates a BTB preview PDF on disk without opening the share sheet. */
export async function generateRunPreviewPdfPath(runId: string): Promise<string> {
  return writeRunExport(runId, 'btb', { persistExport: false });
}

export async function exportSetupPreviewPdf(templateId: string): Promise<string> {
  const [templateRecord, setupModel] = await Promise.all([getTemplate(templateId), getSetupModel(templateId)]);
  if (!templateRecord?.pdfPath) {
    throw new Error('PDF-Vorlage fehlt.');
  }
  if (!setupModel) {
    throw new Error('Setup-Modell fehlt.');
  }

  const base64 = await FileSystem.readAsStringAsync(templateRecord.pdfPath, {
    encoding: FileSystem.EncodingType.Base64
  });
  const pdfBytes = base64ToUint8Array(base64);
  const outputBytes = await buildFinalPdfBytes({
    templateBlob: {
      arrayBuffer: async () => bytesToArrayBuffer(pdfBytes)
    },
    setupModel,
    runValues: {}
  });

  const outDir = `${FileSystem.documentDirectory}bautagebuch/previews/`;
  await FileSystem.makeDirectoryAsync(outDir, { intermediates: true });
  const outPath = `${outDir}setup_preview_${templateId}.pdf`;
  await FileSystem.writeAsStringAsync(outPath, uint8ToBase64(outputBytes), {
    encoding: FileSystem.EncodingType.Base64
  });

  await sharePdf(outPath, 'Setup-Vorschau');
  return outPath;
}

export async function shareCachedExport(exportId: string): Promise<void> {
  const record = await getExport(exportId);
  if (!record) throw new Error('Export nicht gefunden.');
  const info = await FileSystem.getInfoAsync(record.filePath);
  if (!info.exists) {
    await exportRunPdf(record.runId, 'merged');
    return;
  }
  await sharePdf(record.filePath, record.fileName);
}

export async function deleteCachedExport(exportId: string): Promise<void> {
  const record = await getExport(exportId);
  if (!record) return;
  await FileSystem.deleteAsync(record.filePath, { idempotent: true });
  await deleteExport(exportId);
}
