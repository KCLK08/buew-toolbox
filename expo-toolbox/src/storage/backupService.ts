import * as FileSystem from 'expo-file-system/legacy';

import {
  checkpointWal,
  getDatabase,
  isDatabaseWriteInProgress,
  resetDatabaseConnection
} from '../database/sqlite';
import { BACKUP_PREFIX, DB_NAME, MAX_BACKUPS } from '../database/schema/constants';
import { nowIso } from '../lib/ids';
import { ensureStorageLayout } from './fileService';

export type BackupReason =
  | 'photo_added'
  | 'record_deleted'
  | 'status_change'
  | 'app_background'
  | 'manual';

export type BackupInfo = {
  name: string;
  uri: string;
  modifiedAt: number;
  createdAtIso: string;
};

const MIN_BACKUP_INTERVAL_MS = 60_000;

let lastBackupAtMs = 0;
let backupInFlight: Promise<string | null> | null = null;
let deferredReason: BackupReason | null = null;

async function listBackups(backupDir: string): Promise<BackupInfo[]> {
  const names = await FileSystem.readDirectoryAsync(backupDir);
  const backups: BackupInfo[] = [];
  for (const name of names) {
    if (!name.startsWith(BACKUP_PREFIX) || !name.endsWith('.db')) continue;
    const uri = `${backupDir}${name}`;
    const info = await FileSystem.getInfoAsync(uri);
    if (!info.exists) continue;
    const modifiedAt = info.modificationTime ? info.modificationTime * 1000 : 0;
    backups.push({
      name,
      uri,
      modifiedAt,
      createdAtIso: new Date(modifiedAt || Date.now()).toISOString()
    });
  }
  return backups.sort((a, b) => b.modifiedAt - a.modifiedAt);
}

function sqliteDatabaseUri(): string {
  const root = FileSystem.documentDirectory;
  if (!root) {
    throw new Error('documentDirectory ist nicht verfügbar.');
  }
  return `${root}SQLite/${DB_NAME}`;
}

async function rotateBackups(backupDir: string): Promise<void> {
  const existing = await listBackups(backupDir);
  const obsolete = existing.slice(MAX_BACKUPS);
  for (const item of obsolete) {
    await FileSystem.deleteAsync(item.uri, { idempotent: true });
  }
}

async function createSafeDatabaseBackup(): Promise<string | null> {
  if (isDatabaseWriteInProgress()) {
    return null;
  }

  // Ensure DB handle exists, then flush WAL to the main DB file.
  await getDatabase();
  await checkpointWal();

  if (isDatabaseWriteInProgress()) {
    return null;
  }

  const { backups } = await ensureStorageLayout();
  const source = sqliteDatabaseUri();
  const sourceInfo = await FileSystem.getInfoAsync(source);
  if (!sourceInfo.exists) {
    return null;
  }

  const stamp = nowIso().replace(/[:.]/g, '-');
  const targetName = `${BACKUP_PREFIX}${stamp}.db`;
  const target = `${backups}${targetName}`;
  await FileSystem.copyAsync({ from: source, to: target });
  await rotateBackups(backups);
  lastBackupAtMs = Date.now();
  return target;
}

/**
 * Request a throttled, write-safe backup.
 * Skips when a write is in progress or when called more than once per minute
 * (except manual).
 */
export async function requestDatabaseBackup(reason: BackupReason): Promise<string | null> {
  if (isDatabaseWriteInProgress()) {
    deferredReason = reason;
    return null;
  }

  const now = Date.now();
  if (reason !== 'manual' && now - lastBackupAtMs < MIN_BACKUP_INTERVAL_MS) {
    return null;
  }

  if (backupInFlight) {
    deferredReason = reason;
    return backupInFlight;
  }

  backupInFlight = createSafeDatabaseBackup()
    .catch(() => null)
    .finally(() => {
      backupInFlight = null;
      const pending = deferredReason;
      deferredReason = null;
      if (pending && !isDatabaseWriteInProgress()) {
        void requestDatabaseBackup(pending);
      }
    });

  return backupInFlight;
}

/** Flush any backup deferred because a write was in progress. */
export async function flushDeferredBackup(): Promise<string | null> {
  if (!deferredReason) return null;
  const reason = deferredReason;
  deferredReason = null;
  return requestDatabaseBackup(reason);
}

export async function getLatestBackupInfo(): Promise<BackupInfo | null> {
  const { backups } = await ensureStorageLayout();
  const existing = await listBackups(backups);
  return existing[0] ?? null;
}

export async function restoreDatabaseFromBackup(backupUri: string): Promise<boolean> {
  if (!backupUri) return false;
  if (isDatabaseWriteInProgress()) {
    throw new Error('Wiederherstellung während laufendem Schreibvorgang nicht möglich.');
  }

  await resetDatabaseConnection();
  const target = sqliteDatabaseUri();
  const parent = target.replace(/\/[^/]+$/, '/');
  const parentInfo = await FileSystem.getInfoAsync(parent);
  if (!parentInfo.exists) {
    await FileSystem.makeDirectoryAsync(parent, { intermediates: true });
  }
  await FileSystem.copyAsync({ from: backupUri, to: target });
  return true;
}

export async function listDatabaseBackups(): Promise<BackupInfo[]> {
  const { backups } = await ensureStorageLayout();
  return listBackups(backups);
}

/** @deprecated Use requestDatabaseBackup with an explicit reason. */
export async function createDatabaseBackup(): Promise<string | null> {
  return requestDatabaseBackup('manual');
}

/** @deprecated Use restoreDatabaseFromBackup after user confirmation. */
export async function restoreDatabaseFromLatestBackup(): Promise<boolean> {
  const latest = await getLatestBackupInfo();
  if (!latest) return false;
  return restoreDatabaseFromBackup(latest.uri);
}
