const MAX_BACKUPS = 3;

function nowIso() {
  return new Date().toISOString();
}

function createBackupId() {
  return `bak_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

/**
 * Rotate IndexedDB snapshots for crash recovery.
 * Keeps at most MAX_BACKUPS entries in `db_backups`.
 */
export async function createIndexedDbBackup(db, { label = 'auto' } = {}) {
  if (!db?.table) {
    throw new Error('Backup benötigt eine Dexie-Datenbank.');
  }

  const tableNames = db.tables.map((table) => table.name).filter((name) => name !== 'db_backups');
  const snapshot = {};

  for (const name of tableNames) {
    snapshot[name] = await db.table(name).toArray();
  }

  const record = {
    id: createBackupId(),
    label: String(label || 'auto'),
    createdAt: nowIso(),
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

  return record.id;
}

export async function listIndexedDbBackups(db) {
  if (!db?.db_backups) return [];
  return db.db_backups.orderBy('createdAt').reverse().toArray();
}

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
      await table.clear();
      const rows = Array.isArray(backup.snapshot[name]) ? backup.snapshot[name] : [];
      if (rows.length > 0) {
        await table.bulkPut(rows);
      }
    }
  });

  return true;
}

export async function restoreLatestIndexedDbBackup(db) {
  const backups = await listIndexedDbBackups(db);
  if (backups.length === 0) return false;
  await restoreIndexedDbBackup(db, backups[0].id);
  return true;
}
