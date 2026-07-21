import * as FileSystem from 'expo-file-system/legacy';
import * as Sharing from 'expo-sharing';

import { uint8ToBase64 } from '../../../lib/binary';
import type { SiteReportExport, SiteReportProtocol } from '../db/database';
import {
  deleteExport,
  getExport,
  getExportByProtocol,
  getProtocol,
  loadLogo,
  upsertExportByProtocol
} from '../db/database';
import { guessImageExtension } from '../lib/native-image';
import { exportToPdfData } from '../lib/pdf.js';
import { exportToXlsxData } from '../lib/xlsx-export.js';

const EXPORT_DIR = `${FileSystem.documentDirectory}sitereport/exports/`;

async function prepareEntries(protocol: SiteReportProtocol) {
  return Promise.all(
    protocol.entries.map(async (entry) => {
      if (!entry.photoPath) {
        return { ...entry, photoBase64: null };
      }
      const photoBase64 = await FileSystem.readAsStringAsync(entry.photoPath, {
        encoding: FileSystem.EncodingType.Base64
      });
      return {
        ...entry,
        photoBase64,
        photoExtension: guessImageExtension(entry.photoPath),
        photoMimeType: 'image/jpeg'
      };
    })
  );
}

async function exportPayload(protocol: SiteReportProtocol, entries: Awaited<ReturnType<typeof prepareEntries>>) {
  const logoDataUrl = await loadLogo();
  return {
    protocolTitle: protocol.protocolTitle,
    projectName: protocol.projectName,
    protocolDate: protocol.protocolDate,
    protocolDescription: protocol.protocolDescription,
    attendees: protocol.attendees,
    logoDataUrl,
    columns: protocol.columns,
    entries
  };
}

async function ensureExportDir() {
  await FileSystem.makeDirectoryAsync(EXPORT_DIR, { intermediates: true });
}

async function writeExportFile(path: string, bytes: Uint8Array) {
  await FileSystem.writeAsStringAsync(path, uint8ToBase64(bytes), {
    encoding: FileSystem.EncodingType.Base64
  });
  return path;
}

function exportMeta(protocol: SiteReportProtocol) {
  return {
    protocolId: protocol.id,
    protocolTitle: protocol.protocolTitle,
    projectName: protocol.projectName,
    protocolDate: protocol.protocolDate
  };
}

async function shareFile(path: string, mimeType: string, title: string) {
  if (await Sharing.isAvailableAsync()) {
    await Sharing.shareAsync(path, { mimeType, dialogTitle: title });
  }
  return path;
}

export async function exportProtocolPdf(protocol: SiteReportProtocol): Promise<string> {
  await ensureExportDir();
  const entries = await prepareEntries(protocol);
  const result = await exportToPdfData(await exportPayload(protocol, entries));
  const outPath = `${EXPORT_DIR}${result.filename}`;
  await writeExportFile(outPath, result.bytes);
  await upsertExportByProtocol({
    ...exportMeta(protocol),
    pdfPath: outPath,
    pdfFilename: result.filename
  });
  return shareFile(outPath, 'application/pdf', protocol.protocolTitle);
}

export async function exportProtocolXlsx(protocol: SiteReportProtocol): Promise<string> {
  await ensureExportDir();
  const entries = await prepareEntries(protocol);
  const result = await exportToXlsxData(await exportPayload(protocol, entries));
  const outPath = `${EXPORT_DIR}${result.filename}`;
  const bytes = result.buffer instanceof ArrayBuffer ? new Uint8Array(result.buffer) : new Uint8Array(result.buffer);
  await writeExportFile(outPath, bytes);
  await upsertExportByProtocol({
    ...exportMeta(protocol),
    xlsxPath: outPath,
    xlsxFilename: result.filename
  });
  return shareFile(
    outPath,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    protocol.protocolTitle
  );
}

async function ensurePdfExport(protocol: SiteReportProtocol): Promise<SiteReportExport> {
  const existing = await getExportByProtocol(protocol.id);
  if (existing?.pdfPath) {
    const info = await FileSystem.getInfoAsync(existing.pdfPath);
    if (info.exists) return existing;
  }
  await exportProtocolPdf(protocol);
  const record = await getExportByProtocol(protocol.id);
  if (!record?.pdfPath) throw new Error('PDF-Export konnte nicht erstellt werden.');
  return record;
}

async function ensureXlsxExport(protocol: SiteReportProtocol): Promise<SiteReportExport> {
  const existing = await getExportByProtocol(protocol.id);
  if (existing?.xlsxPath) {
    const info = await FileSystem.getInfoAsync(existing.xlsxPath);
    if (info.exists) return existing;
  }
  await exportProtocolXlsx(protocol);
  const record = await getExportByProtocol(protocol.id);
  if (!record?.xlsxPath) throw new Error('XLSX-Export konnte nicht erstellt werden.');
  return record;
}

export async function shareCachedExport(exportId: string, format: 'pdf' | 'xlsx'): Promise<void> {
  let record = await getExport(exportId);
  if (!record) throw new Error('Export nicht gefunden.');

  const path = format === 'pdf' ? record.pdfPath : record.xlsxPath;
  if (path) {
    const info = await FileSystem.getInfoAsync(path);
    if (info.exists) {
      await shareFile(
        path,
        format === 'pdf'
          ? 'application/pdf'
          : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        record.protocolTitle || record.projectName
      );
      return;
    }
  }

  const protocol = await getProtocol(record.protocolId);
  if (!protocol) throw new Error('Protokoll nicht gefunden.');
  record = format === 'pdf' ? await ensurePdfExport(protocol) : await ensureXlsxExport(protocol);
  const nextPath = format === 'pdf' ? record.pdfPath : record.xlsxPath;
  if (!nextPath) throw new Error('Export nicht verfügbar.');
  await shareFile(
    nextPath,
    format === 'pdf'
      ? 'application/pdf'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    record.protocolTitle || record.projectName
  );
}

export async function deleteCachedExport(exportId: string): Promise<void> {
  const record = await getExport(exportId);
  if (!record) return;
  for (const path of [record.pdfPath, record.xlsxPath]) {
    if (!path) continue;
    await FileSystem.deleteAsync(path, { idempotent: true });
  }
  await deleteExport(exportId);
}