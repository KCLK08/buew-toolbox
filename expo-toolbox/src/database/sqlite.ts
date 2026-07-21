import * as SQLite from 'expo-sqlite';

import { migrations } from './migrations';
import { DB_NAME, SCHEMA_VERSION } from './schema/constants';
import { nowIso } from '../lib/ids';

let dbPromise: Promise<SQLite.SQLiteDatabase> | null = null;

export async function getDatabase(): Promise<SQLite.SQLiteDatabase> {
  if (!dbPromise) {
    dbPromise = SQLite.openDatabaseAsync(DB_NAME).catch((error) => {
      dbPromise = null;
      throw error;
    });
  }
  return dbPromise;
}

export async function withTransaction<T>(work: (db: SQLite.SQLiteDatabase) => Promise<T>): Promise<T> {
  const db = await getDatabase();
  let result!: T;
  await db.withTransactionAsync(async () => {
    result = await work(db);
  });
  return result;
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
  const applied = await getAppliedVersions(db);
  const pending = migrations.filter((migration) => !applied.has(migration.version)).sort((a, b) => a.version - b.version);
  const from = applied.size > 0 ? Math.max(...applied) : 0;

  for (const migration of pending) {
    await db.withTransactionAsync(async () => {
      for (const statement of migration.up) {
        await db.execAsync(statement);
      }
      await db.runAsync(
        'INSERT INTO schema_migrations (version, name, applied_at) VALUES (?, ?, ?)',
        migration.version,
        migration.name,
        nowIso()
      );
      await db.runAsync(
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
