import * as FileSystem from 'expo-file-system/legacy';
import * as Sharing from 'expo-sharing';

import { uint8ToBase64 } from '../../../lib/binary';
import type { SiteReportProtocol } from '../db/database';
import { guessImageExtension } from '../lib/native-image';
import { exportToPdfData } from '../lib/pdf.js';
import { exportToXlsxData } from '../lib/xlsx-export.js';

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

function exportPayload(protocol: SiteReportProtocol, entries: Awaited<ReturnType<typeof prepareEntries>>) {
  return {
    protocolTitle: protocol.protocolTitle,
    projectName: protocol.projectName,
    protocolDate: protocol.protocolDate,
    protocolDescription: protocol.protocolDescription,
    attendees: protocol.attendees,
    logoDataUrl: '',
    columns: protocol.columns,
    entries
  };
}

async function writeAndShare(path: string, bytes: Uint8Array, mimeType: string, title: string) {
  await FileSystem.writeAsStringAsync(path, uint8ToBase64(bytes), {
    encoding: FileSystem.EncodingType.Base64
  });
  if (await Sharing.isAvailableAsync()) {
    await Sharing.shareAsync(path, { mimeType, dialogTitle: title });
  }
  return path;
}

export async function exportProtocolPdf(protocol: SiteReportProtocol): Promise<string> {
  const entries = await prepareEntries(protocol);
  const result = await exportToPdfData(exportPayload(protocol, entries));
  const outDir = `${FileSystem.documentDirectory}sitereport/exports/`;
  await FileSystem.makeDirectoryAsync(outDir, { intermediates: true });
  const outPath = `${outDir}${result.filename}`;
  return writeAndShare(outPath, result.bytes, 'application/pdf', protocol.protocolTitle);
}

export async function exportProtocolXlsx(protocol: SiteReportProtocol): Promise<string> {
  const entries = await prepareEntries(protocol);
  const result = await exportToXlsxData(exportPayload(protocol, entries));
  const outDir = `${FileSystem.documentDirectory}sitereport/exports/`;
  await FileSystem.makeDirectoryAsync(outDir, { intermediates: true });
  const outPath = `${outDir}${result.filename}`;
  const bytes = result.buffer instanceof ArrayBuffer ? new Uint8Array(result.buffer) : new Uint8Array(result.buffer);
  return writeAndShare(
    outPath,
    bytes,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    protocol.protocolTitle
  );
}
