/**
 * Soft-delete purge preparation only.
 * Builds a dry-run plan for later cleanup. Does not delete anything.
 */

export const SOFT_DELETE_RETENTION_DAYS = 30;

/**
 * @param {import('dexie').Dexie} db
 * @param {{ retentionDays?: number }} [options]
 * @returns {Promise<{
 *   retentionDays: number,
 *   cutoffIso: string,
 *   candidates: { table: string, id: string, deletedAt: string }[],
 *   purgeEnabled: false
 * }>}
 */
export async function planSoftDeletePurge(db, options = {}) {
  const retentionDays = options.retentionDays ?? SOFT_DELETE_RETENTION_DAYS;
  const cutoff = Date.now() - retentionDays * 24 * 60 * 60 * 1000;
  const cutoffIso = new Date(cutoff).toISOString();

  /** @type {{ table: string, idKey: string }[]} */
  const tables = [
    { table: 'templates', idKey: 'templateId' },
    { table: 'runs', idKey: 'runId' },
    { table: 'exports', idKey: 'exportId' },
    { table: 'photo_assets', idKey: 'id' }
  ];

  /** @type {{ table: string, id: string, deletedAt: string }[]} */
  const candidates = [];

  for (const { table, idKey } of tables) {
    if (!db.table(table)) continue;
    const rows = await db.table(table).toArray();
    for (const row of rows) {
      const deletedAt = String(row?.deleted_at || '').trim();
      if (!deletedAt) continue;
      const deletedMs = Date.parse(deletedAt);
      if (Number.isFinite(deletedMs) && deletedMs <= cutoff) {
        candidates.push({
          table,
          id: String(row?.[idKey] || row?.id || ''),
          deletedAt
        });
      }
    }
  }

  const plan = {
    retentionDays,
    cutoffIso,
    candidates,
    purgeEnabled: false
  };

  if (candidates.length > 0) {
    console.info('[soft-delete-purge] Dry-run plan (purge disabled):', plan);
  }

  return plan;
}

/**
 * Intentionally disabled. Hard purge must be enabled in a later phase.
 * @throws {Error}
 */
export async function executeSoftDeletePurge() {
  throw new Error(
    'Soft-delete purge is prepared but not enabled. No records were deleted.'
  );
}
