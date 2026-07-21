import * as FileSystem from 'expo-file-system/legacy';

import {
  checkpointWal,
  getDatabase,
  isDatabaseWriteInProgress,
  resetDatabaseConnection
} from '../database/sqlite';
import {
  BACKUP_PREFIX,
  BAUTAGEBUCH_BACKUP_PREFIX,
  BAUTAGEBUCH_DB_NAME,
  DB_NAME,
  MAX_BACKUPS,
  SITEREPORT_BACKUP_PREFIX,
  SITEREPORT_DB_NAME
} from '../database/schema/constants';
import { resetBautagebuchDatabaseConnection } from '../native/bautagebuch/db/database';
import { resetSiteReportDatabaseConnection } from '../native/sitereport/db/database';
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
  stamp: string;
};

const MIN_BACKUP_INTERVAL_MS = 60_000;

const DATABASE_BACKUP_TARGETS = [
  { dbName: DB_NAME, prefix: BACKUP_PREFIX },
  { dbName: SITEREPORT_DB_NAME, prefix: SITEREPORT_BACKUP_PREFIX },
  { dbName: BAUTAGEBUCH_DB_NAME, prefix: BAUTAGEBUCH_BACKUP_PREFIX }
] as const;

let lastBackupAtMs = 0;
let backupInFlight: Promise<string | null> | null = null;
let deferredReason: BackupReason | null = null;

function sqliteDatabaseUri(dbName: string): string {
  const root = FileSystem.documentDirectory;
  if (!root) {
    throw new Error('documentDirectory ist nicht verfügbar.');
  }
  return `${root}SQLite/${dbName}`;
}

function extractBackupStamp(name: string): string | null {
  for (const target of DATABASE_BACKUP_TARGETS) {
    if (!name.startsWith(target.prefix) || !name.endsWith('.db')) continue;
    return name.slice(target.prefix.length, -'.db'.length);
  }
  return null;
}

async function listPrimaryBackups(backupDir: string): Promise<BackupInfo[]> {
  const names = await FileSystem.readDirectoryAsync(backupDir);
  const backups: BackupInfo[] = [];
  for (const name of names) {
    if (!name.startsWith(BACKUP_PREFIX) || !name.endsWith('.db')) continue;
    const stamp = extractBackupStamp(name);
    if (!stamp) continue;
    const uri = `${backupDir}${name}`;
    const info = await FileSystem.getInfoAsync(uri);
    if (!info.exists) continue;
    const modifiedAt = info.modificationTime ? info.modificationTime * 1000 : 0;
    backups.push({
      name,
      uri,
      modifiedAt,
      createdAtIso: new Date(modifiedAt || Date.now()).toISOString(),
      stamp
    });
  }
  return backups.sort((a, b) => b.modifiedAt - a.modifiedAt);
}

async function rotateBackups(backupDir: string): Promise<void> {
  const existing = await listPrimaryBackups(backupDir);
  const obsoleteStamps = existing.slice(MAX_BACKUPS).map((item) => item.stamp);
  for (const stamp of obsoleteStamps) {
    for (const target of DATABASE_BACKUP_TARGETS) {
      const uri = `${backupDir}${target.prefix}${stamp}.db`;
      await FileSystem.deleteAsync(uri, { idempotent: true });
    }
  }
}

async function copyDatabaseBackup(stamp: string, backupDir: string): Promise<string | null> {
  let primaryTarget: string | null = null;
  for (const target of DATABASE_BACKUP_TARGETS) {
    const source = sqliteDatabaseUri(target.dbName);
    const sourceInfo = await FileSystem.getInfoAsync(source);
    if (!sourceInfo.exists) continue;
    const backupName = `${target.prefix}${stamp}.db`;
    const backupUri = `${backupDir}${backupName}`;
    await FileSystem.copyAsync({ from: source, to: backupUri });
    if (target.dbName === DB_NAME) {
      primaryTarget = backupUri;
    }
  }
  return primaryTarget;
}

async function createSafeDatabaseBackup(): Promise<string | null> {
  if (isDatabaseWriteInProgress()) {
    return null;
  }

  await getDatabase();
  await checkpointWal();

  if (isDatabaseWriteInProgress()) {
    return null;
  }

  const { backups } = await ensureStorageLayout();
  const primarySource = sqliteDatabaseUri(DB_NAME);
  const primaryInfo = await FileSystem.getInfoAsync(primarySource);
  if (!primaryInfo.exists) {
    return null;
  }

  const stamp = nowIso().replace(/[:.]/g, '-');
  const primaryTarget = await copyDatabaseBackup(stamp, backups);
  if (!primaryTarget) {
    return null;
  }
  await rotateBackups(backups);
  lastBackupAtMs = Date.now();
  return primaryTarget;
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
  const existing = await listPrimaryBackups(backups);
  return existing[0] ?? null;
}

async function restoreDatabaseFile(dbName: string, backupUri: string): Promise<void> {
  const target = sqliteDatabaseUri(dbName);
  const parent = target.replace(/\/[^/]+$/, '/');
  const parentInfo = await FileSystem.getInfoAsync(parent);
  if (!parentInfo.exists) {
    await FileSystem.makeDirectoryAsync(parent, { intermediates: true });
  }
  await FileSystem.copyAsync({ from: backupUri, to: target });
}

export async function restoreDatabaseFromBackup(backupUri: string): Promise<boolean> {
  if (!backupUri) return false;
  if (isDatabaseWriteInProgress()) {
    throw new Error('Wiederherstellung während laufendem Schreibvorgang nicht möglich.');
  }

  const stamp = extractBackupStamp(backupUri.split('/').pop() || '');
  if (!stamp) return false;

  const { backups } = await ensureStorageLayout();
  await Promise.all([
    resetDatabaseConnection(),
    resetSiteReportDatabaseConnection(),
    resetBautagebuchDatabaseConnection()
  ]);

  let restoredAny = false;
  for (const target of DATABASE_BACKUP_TARGETS) {
    const source = `${backups}${target.prefix}${stamp}.db`;
    const info = await FileSystem.getInfoAsync(source);
    if (!info.exists) continue;
    await restoreDatabaseFile(target.dbName, source);
    restoredAny = true;
  }

  return restoredAny;
}

export async function listDatabaseBackups(): Promise<BackupInfo[]> {
  const { backups } = await ensureStorageLayout();
  return listPrimaryBackups(backups);
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
