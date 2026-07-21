import * as FileSystem from 'expo-file-system/legacy';

import { BACKUP_DIR, DOCUMENT_DIR, PHOTO_DIR } from '../database/schema/constants';
import { createUuid, nowIso } from '../lib/ids';

function requireDocumentDirectory(): string {
  const root = FileSystem.documentDirectory;
  if (!root) {
    throw new Error('documentDirectory ist nicht verfügbar.');
  }
  return root;
}

async function ensureDir(path: string): Promise<void> {
  const info = await FileSystem.getInfoAsync(path);
  if (!info.exists) {
    await FileSystem.makeDirectoryAsync(path, { intermediates: true });
  }
}

export async function ensureStorageLayout(): Promise<{
  photos: string;
  documents: string;
  backups: string;
}> {
  const root = requireDocumentDirectory();
  const photos = `${root}${PHOTO_DIR}/`;
  const documents = `${root}${DOCUMENT_DIR}/`;
  const backups = `${root}${BACKUP_DIR}/`;
  await Promise.all([ensureDir(photos), ensureDir(documents), ensureDir(backups)]);
  return { photos, documents, backups };
}

export type StoredFile = {
  id: string;
  uri: string;
  relativePath: string;
  byteSize: number;
  mimeType: string;
  createdAt: string;
};

function extensionForMime(mimeType: string): string {
  if (mimeType.includes('png')) return 'png';
  if (mimeType.includes('webp')) return 'webp';
  if (mimeType.includes('heic')) return 'heic';
  if (mimeType.includes('pdf')) return 'pdf';
  return 'jpg';
}

/** Persist a local file into documentDirectory (never cacheDirectory). */
export async function persistLocalFile(options: {
  sourceUri: string;
  mimeType?: string;
  kind?: 'photo' | 'document';
  id?: string;
}): Promise<StoredFile> {
  const layout = await ensureStorageLayout();
  const id = options.id ?? (await createUuid());
  const mimeType = options.mimeType ?? 'image/jpeg';
  const kind = options.kind ?? 'photo';
  const folder = kind === 'document' ? layout.documents : layout.photos;
  const relativePath = `${kind === 'document' ? DOCUMENT_DIR : PHOTO_DIR}/${id}.${extensionForMime(mimeType)}`;
  const targetUri = `${folder}${id}.${extensionForMime(mimeType)}`;

  await FileSystem.copyAsync({
    from: options.sourceUri,
    to: targetUri
  });

  const info = await FileSystem.getInfoAsync(targetUri);
  const byteSize = info.exists && 'size' in info && typeof info.size === 'number' ? info.size : 0;

  return {
    id,
    uri: targetUri,
    relativePath,
    byteSize,
    mimeType,
    createdAt: nowIso()
  };
}

export async function fileExists(uri: string): Promise<boolean> {
  const info = await FileSystem.getInfoAsync(uri);
  return info.exists;
}

export async function deleteFileIfExists(uri: string): Promise<void> {
  const info = await FileSystem.getInfoAsync(uri);
  if (info.exists) {
    await FileSystem.deleteAsync(uri, { idempotent: true });
  }
}

export function resolveDocumentUri(relativePath: string): string {
  const root = requireDocumentDirectory();
  if (relativePath.startsWith(root)) return relativePath;
  return `${root}${relativePath.replace(/^\//, '')}`;
}

export { FileSystem };
