/**
 * Orphan cleanup (report-only).
 * Finds photo_assets that are not referenced by any active run photoDoc entry.
 * Never deletes anything automatically.
 */

import { photoAssetKey } from './photo-storage.js';

/**
 * @param {import('dexie').Dexie} db
 * @returns {Promise<{
 *   scannedAssets: number,
 *   orphanAssetIds: string[],
 *   reportedAt: string
 * }>}
 */
export async function findOrphanPhotoAssets(db) {
  const [assets, runs] = await Promise.all([db.photo_assets.toArray(), db.runs.toArray()]);

  const referenced = new Set();
  for (const run of runs) {
    if (String(run?.deleted_at || '').trim()) continue;
    const runId = String(run?.runId || '').trim();
    if (!runId) continue;
    const entries = Array.isArray(run?.photoDoc?.entries) ? run.photoDoc.entries : [];
    for (const entry of entries) {
      const entryId = String(entry?.id || '').trim();
      if (!entryId) continue;
      referenced.add(photoAssetKey(runId, entryId));
    }
  }

  const orphanAssetIds = assets
    .filter((asset) => {
      if (String(asset?.deleted_at || '').trim()) return false;
      const id = String(asset?.id || '').trim();
      if (!id) return false;
      return !referenced.has(id);
    })
    .map((asset) => String(asset.id));

  const report = {
    scannedAssets: assets.length,
    orphanAssetIds,
    reportedAt: new Date().toISOString()
  };

  if (orphanAssetIds.length > 0) {
    console.warn('[orphan-cleanup] Unreferenced photo assets (not deleted):', report);
  }

  return report;
}
