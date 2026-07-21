import * as SQLite from 'expo-sqlite';

import { migrations } from './migrations';
import { DB_NAME, SCHEMA_VERSION } from './schema/constants';
import { nowIso } from '../lib/ids';

let dbPromise: Promise<SQLite.SQLiteDatabase> | null = null;
let activeWriteCount = 0;

export function isDatabaseWriteInProgress(): boolean {
  return activeWriteCount > 0;
}

export function markDatabaseWriteStarted(): void {
  activeWriteCount += 1;
}

export function markDatabaseWriteFinished(): void {
  activeWriteCount = Math.max(0, activeWriteCount - 1);
}

export async function getDatabase(): Promise<SQLite.SQLiteDatabase> {
  if (!dbPromise) {
    dbPromise = (async () => {
      const db = await SQLite.openDatabaseAsync(DB_NAME);
      await db.execAsync('PRAGMA foreign_keys = ON;');
      await db.execAsync('PRAGMA journal_mode = WAL;');
      await db.execAsync('PRAGMA synchronous = NORMAL;');
      return db;
    })().catch((error) => {
      dbPromise = null;
      throw error;
    });
  }
  return dbPromise;
}

export async function withTransaction<T>(work: (db: SQLite.SQLiteDatabase) => Promise<T>): Promise<T> {
  const db = await getDatabase();
  markDatabaseWriteStarted();
  try {
    let result!: T;
    await db.withTransactionAsync(async () => {
      result = await work(db);
    });
    return result;
  } finally {
    markDatabaseWriteFinished();
  }
}

async function getAppliedVersions(db: SQLite.SQLiteDatabase): Promise<Set<number>> {
  await db.execAsync(`
    CREATE TABLE IF NOT EXISTS schema_migrations (
      version INTEGER PRIMARY KEY NOT NULL,
      name TEXT NOT NULL,
      applied_at TEXT NOT NULL
    );
  `);
  const rows = await db.getAllAsync<{ version: number }>('SELECT version FROM schema_migrations');
  return new Set(rows.map((row) => Number(row.version)));
}

export async function runMigrations(): Promise<{ from: number; to: number }> {
  const db = await getDatabase();
  await db.execAsync('PRAGMA foreign_keys = ON;');
  await db.execAsync('PRAGMA journal_mode = WAL;');
  await db.execAsync('PRAGMA synchronous = NORMAL;');
  const applied = await getAppliedVersions(db);
  const pending = migrations.filter((migration) => !applied.has(migration.version)).sort((a, b) => a.version - b.version);
  const from = applied.size > 0 ? Math.max(...applied) : 0;

  for (const migration of pending) {
    await withTransaction(async (txDb) => {
      for (const statement of migration.up) {
        await txDb.execAsync(statement);
      }
      await txDb.runAsync(
        'INSERT INTO schema_migrations (version, name, applied_at) VALUES (?, ?, ?)',
        migration.version,
        migration.name,
        nowIso()
      );
      await txDb.runAsync(
        `INSERT INTO app_meta (key, value, updated_at) VALUES ('schema_version', ?, ?)
         ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at`,
        String(migration.version),
        nowIso()
      );
    });
  }

  return { from, to: SCHEMA_VERSION };
}

export async function tableExists(tableName: string): Promise<boolean> {
  const db = await getDatabase();
  const row = await db.getFirstAsync<{ name: string }>(
    `SELECT name FROM sqlite_master WHERE type = 'table' AND name = ?`,
    tableName
  );
  return Boolean(row?.name);
}

export async function resetDatabaseConnection(): Promise<void> {
  if (!dbPromise) return;
  const db = await dbPromise;
  await db.closeAsync();
  dbPromise = null;
}

/** Flush WAL before creating a consistent on-disk copy. */
export async function checkpointWal(): Promise<void> {
  const db = await getDatabase();
  await db.execAsync('PRAGMA wal_checkpoint(FULL);');
}
