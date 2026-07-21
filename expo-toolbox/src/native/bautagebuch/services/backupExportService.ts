import * as FileSystem from 'expo-file-system/legacy';
import * as DocumentPicker from 'expo-document-picker';
import * as Sharing from 'expo-sharing';
import JSZip from 'jszip';

import { BAUTAGEBUCH_DB_NAME } from '../../../database/schema/constants';
import {
  getBautagebuchStorageRoot,
  initBautagebuchDatabase,
  listTemplates,
  resetBautagebuchDatabaseConnection
} from '../db/database';
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

async function writeBase64File(path: string, base64: string): Promise<void> {
  const parent = path.replace(/\/[^/]+$/, '/');
  await FileSystem.makeDirectoryAsync(parent, { intermediates: true });
  await FileSystem.writeAsStringAsync(path, base64, {
    encoding: FileSystem.EncodingType.Base64
  });
}

async function restoreZipPrefix(zip: JSZip, prefix: string, targetDir: string): Promise<number> {
  await FileSystem.deleteAsync(targetDir, { idempotent: true });
  await FileSystem.makeDirectoryAsync(targetDir, { intermediates: true });

  let count = 0;
  for (const path of Object.keys(zip.files)) {
    const entry = zip.files[path];
    if (!entry || entry.dir || !path.startsWith(prefix)) continue;
    const relative = path.slice(prefix.length);
    if (!relative) continue;
    const base64 = await entry.async('base64');
    await writeBase64File(`${targetDir}${relative}`, base64);
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

/**
 * Restores a BTB backup ZIP created by exportBautagebuchBackupZip.
 * Replaces SQLite DB and bautagebuch/ file trees (templates, photos, exports).
 */
export async function importBautagebuchBackupZip(fileUri: string): Promise<{
  photoFileCount: number;
  exportFileCount: number;
}> {
  const base64 = await FileSystem.readAsStringAsync(fileUri, {
    encoding: FileSystem.EncodingType.Base64
  });
  const zip = await JSZip.loadAsync(base64, { base64: true });

  const manifestRaw = await zip.file('manifest.json')?.async('string');
  if (!manifestRaw) {
    throw new Error('Ungültiges Backup: manifest.json fehlt.');
  }
  const manifest = JSON.parse(manifestRaw) as { kind?: string };
  if (manifest.kind !== 'bautagebuch-backup') {
    throw new Error('Die Datei ist kein Bautagebuch-Backup.');
  }

  await resetBautagebuchDatabaseConnection();

  const dbEntry = zip.file(`sqlite/${BAUTAGEBUCH_DB_NAME}`);
  if (!dbEntry) {
    throw new Error('Backup enthält keine Bautagebuch-Datenbank.');
  }
  const dbBase64 = await dbEntry.async('base64');
  const dbPath = `${FileSystem.documentDirectory}SQLite/${BAUTAGEBUCH_DB_NAME}`;
  await writeBase64File(dbPath, dbBase64);

  const root = getBautagebuchStorageRoot();
  await restoreZipPrefix(zip, 'templates/', `${root}templates/`);
  const photoFileCount = await restoreZipPrefix(zip, 'photos/', `${root}photos/`);
  const exportFileCount = await restoreZipPrefix(zip, 'exports/', `${root}exports/`);

  await initBautagebuchDatabase();

  return { photoFileCount, exportFileCount };
}

export async function pickAndRestoreBautagebuchBackup(): Promise<{
  photoFileCount: number;
  exportFileCount: number;
} | null> {
  const result = await DocumentPicker.getDocumentAsync({
    type: ['application/zip', 'application/x-zip-compressed', 'application/octet-stream'],
    copyToCacheDirectory: true,
    multiple: false
  });
  if (result.canceled || !result.assets?.[0]?.uri) {
    return null;
  }
  return importBautagebuchBackupZip(result.assets[0].uri);
}
