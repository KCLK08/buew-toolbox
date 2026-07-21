import * as FileSystem from 'expo-file-system/legacy';

import { BACKUP_PREFIX, DB_NAME, MAX_BACKUPS } from '../database/schema/constants';
import { resetDatabaseConnection } from '../database/sqlite';
import { nowIso } from '../lib/ids';
import { ensureStorageLayout } from './fileService';

type BackupInfo = {
  name: string;
  uri: string;
  modifiedAt: number;
};

async function listBackups(backupDir: string): Promise<BackupInfo[]> {
  const names = await FileSystem.readDirectoryAsync(backupDir);
  const backups: BackupInfo[] = [];
  for (const name of names) {
    if (!name.startsWith(BACKUP_PREFIX) || !name.endsWith('.db')) continue;
    const uri = `${backupDir}${name}`;
    const info = await FileSystem.getInfoAsync(uri);
    if (!info.exists) continue;
    backups.push({
      name,
      uri,
      modifiedAt: info.modificationTime ? info.modificationTime * 1000 : 0
    });
  }
  return backups.sort((a, b) => b.modifiedAt - a.modifiedAt);
}

function sqliteDatabaseUri(): string {
  const root = FileSystem.documentDirectory;
  if (!root) {
    throw new Error('documentDirectory ist nicht verfügbar.');
  }
  // Expo SQLite stores DB files under SQLite/ in documentDirectory on native.
  return `${root}SQLite/${DB_NAME}`;
}

export async function createDatabaseBackup(): Promise<string | null> {
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

  const existing = await listBackups(backups);
  const obsolete = existing.slice(MAX_BACKUPS);
  for (const item of obsolete) {
    await FileSystem.deleteAsync(item.uri, { idempotent: true });
  }

  return target;
}

export async function getLatestBackupUri(): Promise<string | null> {
  const { backups } = await ensureStorageLayout();
  const existing = await listBackups(backups);
  return existing[0]?.uri ?? null;
}

export async function restoreDatabaseFromLatestBackup(): Promise<boolean> {
  const latest = await getLatestBackupUri();
  if (!latest) return false;

  await resetDatabaseConnection();
  const target = sqliteDatabaseUri();
  const parent = target.replace(/\/[^/]+$/, '/');
  const parentInfo = await FileSystem.getInfoAsync(parent);
  if (!parentInfo.exists) {
    await FileSystem.makeDirectoryAsync(parent, { intermediates: true });
  }
  await FileSystem.copyAsync({ from: latest, to: target });
  return true;
}

export async function listDatabaseBackups(): Promise<BackupInfo[]> {
  const { backups } = await ensureStorageLayout();
  return listBackups(backups);
}
