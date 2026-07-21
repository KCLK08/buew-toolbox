import * as FileSystem from 'expo-file-system/legacy';
import * as Sharing from 'expo-sharing';
import JSZip from 'jszip';

import { BAUTAGEBUCH_DB_NAME } from '../../../database/schema/constants';
import { getBautagebuchStorageRoot, listTemplates } from '../db/database';
import { nowIso } from '../../../lib/ids';

async function readFileBase64(path: string): Promise<string | null> {
  const info = await FileSystem.getInfoAsync(path);
  if (!info.exists) return null;
  return FileSystem.readAsStringAsync(path, { encoding: FileSystem.EncodingType.Base64 });
}

async function addDirectoryToZip(
  zip: JSZip,
  directory: string,
  zipPrefix: string
): Promise<number> {
  const info = await FileSystem.getInfoAsync(directory);
  if (!info.exists || !info.isDirectory) return 0;

  let count = 0;
  const entries = await FileSystem.readDirectoryAsync(directory);
  for (const name of entries) {
    const fullPath = `${directory}${name}`;
    const entryInfo = await FileSystem.getInfoAsync(fullPath);
    if (!entryInfo.exists) continue;
    if (entryInfo.isDirectory) {
      count += await addDirectoryToZip(zip, `${fullPath}/`, `${zipPrefix}${name}/`);
      continue;
    }
    const base64 = await readFileBase64(fullPath);
    if (!base64) continue;
    zip.file(`${zipPrefix}${name}`, base64, { base64: true });
    count += 1;
  }
  return count;
}

/**
 * Creates a portable BTB backup ZIP with SQLite DB, templates, and photo files.
 */
export async function exportBautagebuchBackupZip(): Promise<string> {
  const stamp = nowIso().replace(/[:.]/g, '-');
  const zip = new JSZip();
  const root = getBautagebuchStorageRoot();

  const dbPath = `${FileSystem.documentDirectory}SQLite/${BAUTAGEBUCH_DB_NAME}`;
  const dbBase64 = await readFileBase64(dbPath);
  if (!dbBase64) {
    throw new Error('Bautagebuch-Datenbank nicht gefunden.');
  }
  zip.file(`sqlite/${BAUTAGEBUCH_DB_NAME}`, dbBase64, { base64: true });

  const templates = await listTemplates();
  for (const template of templates) {
    if (!template.pdfPath) continue;
    const base64 = await readFileBase64(template.pdfPath);
    if (!base64) continue;
    const fileName = template.fileName || `${template.templateId}.pdf`;
    zip.file(`templates/${fileName}`, base64, { base64: true });
  }

  const photoCount = await addDirectoryToZip(zip, `${root}photos/`, 'photos/');
  const exportCount = await addDirectoryToZip(zip, `${root}exports/`, 'exports/');

  zip.file(
    'manifest.json',
    JSON.stringify(
      {
        kind: 'bautagebuch-backup',
        version: 1,
        createdAt: nowIso(),
        templateCount: templates.length,
        photoFileCount: photoCount,
        exportFileCount: exportCount
      },
      null,
      2
    )
  );

  const zipBase64 = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });
  const outDir = `${root}backups/`;
  await FileSystem.makeDirectoryAsync(outDir, { intermediates: true });
  const outPath = `${outDir}BTB_Backup_${stamp}.zip`;
  await FileSystem.writeAsStringAsync(outPath, zipBase64, {
    encoding: FileSystem.EncodingType.Base64
  });

  if (await Sharing.isAvailableAsync()) {
    await Sharing.shareAsync(outPath, {
      mimeType: 'application/zip',
      dialogTitle: 'Bautagebuch-Backup teilen'
    });
  }

  return outPath;
}
