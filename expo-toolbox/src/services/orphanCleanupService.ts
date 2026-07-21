import * as FileSystem from 'expo-file-system/legacy';

import { PHOTO_DIR, DOCUMENT_DIR } from '../database/schema/constants';
import { photoRepository } from '../repositories';
import { ensureStorageLayout, resolveDocumentUri } from '../storage/fileService';

export type OrphanFileReport = {
  uri: string;
  relativePath: string;
  kind: 'photo' | 'document';
};

/**
 * Find files in documentDirectory that have no active DB photo row.
 * Reports only — never deletes.
 */
export async function findOrphanFiles(): Promise<OrphanFileReport[]> {
  const layout = await ensureStorageLayout();
  const activePaths = new Set(
    (await photoRepository.listActivePaths()).map((path) => resolveDocumentUri(path))
  );

  const orphans: OrphanFileReport[] = [];

  async function scanDir(dir: string, kind: 'photo' | 'document', prefix: string) {
    const names = await FileSystem.readDirectoryAsync(dir);
    for (const name of names) {
      if (name.startsWith('.')) continue;
      const uri = `${dir}${name}`;
      const info = await FileSystem.getInfoAsync(uri);
      if (!info.exists || info.isDirectory) continue;
      if (!activePaths.has(uri)) {
        orphans.push({
          uri,
          relativePath: `${prefix}/${name}`,
          kind
        });
      }
    }
  }

  await scanDir(layout.photos, 'photo', PHOTO_DIR);
  await scanDir(layout.documents, 'document', DOCUMENT_DIR);
  return orphans;
}
