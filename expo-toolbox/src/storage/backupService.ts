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

/** Backup is disabled for now — kept for API compatibility with repositories. */
export async function requestDatabaseBackup(_reason: BackupReason): Promise<string | null> {
  return null;
}

/** Backup is disabled for now — kept for API compatibility with repositories. */
export async function flushDeferredBackup(): Promise<string | null> {
  return null;
}

export async function getLatestBackupInfo(): Promise<BackupInfo | null> {
  return null;
}

export async function restoreDatabaseFromBackup(_backupUri: string): Promise<boolean> {
  return false;
}

export async function listDatabaseBackups(): Promise<BackupInfo[]> {
  return [];
}

/** @deprecated Backup disabled. */
export async function createDatabaseBackup(): Promise<string | null> {
  return null;
}

/** @deprecated Backup disabled. */
export async function restoreDatabaseFromLatestBackup(): Promise<boolean> {
  return false;
}
