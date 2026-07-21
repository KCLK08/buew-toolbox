const MAX_BACKUPS = 3;
const MIN_BACKUP_INTERVAL_MS = 60_000;

/** Labels that justify a backup (not regular autosave). */
const BACKUP_REASONS = new Set([
  'photo_added',
  'run_photo_update',
  'record_deleted',
  'run_soft_delete',
  'status_change',
  'app_background',
  'manual'
]);

let lastBackupAtMs = 0;

function nowIso() {
  return new Date().toISOString();
}

function createBackupId() {
  return `bak_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

/**
 * Snapshot table rows for backup.
 * photo_assets: metadata only — binary `data` is never included.
 */
function sanitizeTableRows(tableName, rows) {
  if (tableName !== 'photo_assets') {
    return Array.isArray(rows) ? rows : [];
  }
  return (Array.isArray(rows) ? rows : []).map((row) => {
    if (!row || typeof row !== 'object') return row;
    const { data, ...meta } = row;
    return {
      ...meta,
      hasBinaryData: Boolean(data),
      sizeBytes: Number(row.sizeBytes || (data?.byteLength ?? 0) || 0)
    };
  });
}

/**
 * Rotate IndexedDB snapshots for crash recovery.
 * Keeps at most MAX_BACKUPS entries in `db_backups`.
 * Photo binaries stay in `photo_assets` and are not duplicated into snapshots.
 */
export async function createIndexedDbBackup(db, { label = 'manual' } = {}) {
  if (!db?.table) {
    throw new Error('Backup benötigt eine Dexie-Datenbank.');
  }

  const reason = String(label || 'manual');
  if (!BACKUP_REASONS.has(reason) && reason !== 'manual') {
    return null;
  }

  const now = Date.now();
  if (reason !== 'manual' && now - lastBackupAtMs < MIN_BACKUP_INTERVAL_MS) {
    return null;
  }

  const tableNames = db.tables.map((table) => table.name).filter((name) => name !== 'db_backups');
  const snapshot = {};

  for (const name of tableNames) {
    const rows = await db.table(name).toArray();
    snapshot[name] = sanitizeTableRows(name, rows);
  }

  const record = {
    id: createBackupId(),
    label: reason,
    createdAt: nowIso(),
    includesPhotoBinaries: false,
    snapshot
  };

  await db.transaction('rw', db.db_backups, async () => {
    await db.db_backups.put(record);
    const all = await db.db_backups.orderBy('createdAt').reverse().toArray();
    const obsolete = all.slice(MAX_BACKUPS);
    for (const entry of obsolete) {
      await db.db_backups.delete(entry.id);
    }
  });

  lastBackupAtMs = now;
  return record.id;
}

export async function listIndexedDbBackups(db) {
  if (!db?.db_backups) return [];
  return db.db_backups.orderBy('createdAt').reverse().toArray();
}

/**
 * Restore a metadata snapshot.
 * For photo_assets: merge metadata from backup while preserving existing binary `data`
 * when the snapshot intentionally omitted it.
 */
export async function restoreIndexedDbBackup(db, backupId) {
  const backup = await db.db_backups.get(backupId);
  if (!backup?.snapshot || typeof backup.snapshot !== 'object') {
    throw new Error('Backup nicht gefunden oder ungültig.');
  }

  const tableNames = Object.keys(backup.snapshot);
  const tables = tableNames.map((name) => db.table(name));

  await db.transaction('rw', ...tables, async () => {
    for (const name of tableNames) {
      const table = db.table(name);
      const rows = Array.isArray(backup.snapshot[name]) ? backup.snapshot[name] : [];

      if (name === 'photo_assets') {
        const existing = await table.toArray();
        const existingById = new Map(existing.map((row) => [String(row.id), row]));
        await table.clear();
        const merged = rows.map((row) => {
          const id = String(row?.id || '');
          const previous = existingById.get(id);
          const { hasBinaryData, ...meta } = row || {};
          return {
            ...meta,
            data: previous?.data ?? row?.data ?? null
          };
        });
        // Keep photo assets that exist only locally (protect binaries not in snapshot).
        for (const [id, row] of existingById.entries()) {
          if (!merged.some((item) => String(item.id) === id)) {
            merged.push(row);
          }
        }
        if (merged.length > 0) {
          await table.bulkPut(merged);
        }
        continue;
      }

      await table.clear();
      if (rows.length > 0) {
        await table.bulkPut(rows);
      }
    }
  });

  return {
    restored: true,
    backupId,
    createdAt: backup.createdAt,
    label: backup.label
  };
}

export async function restoreLatestIndexedDbBackup(db) {
  const backups = await listIndexedDbBackups(db);
  if (backups.length === 0) return null;
  return restoreIndexedDbBackup(db, backups[0].id);
}

export function getLatestBackupSummary(backups) {
  if (!Array.isArray(backups) || backups.length === 0) return null;
  const latest = backups[0];
  return {
    id: latest.id,
    createdAt: latest.createdAt,
    label: latest.label
  };
}
