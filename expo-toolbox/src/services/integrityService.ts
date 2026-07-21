import { TABLES } from '../database/schema/constants';
import { getDatabase, runMigrations, tableExists } from '../database/sqlite';
import { getLatestBackupInfo, restoreDatabaseFromBackup } from '../storage/backupService';
import { ensureStorageLayout, fileExists, resolveDocumentUri } from '../storage/fileService';
import { findOrphanFiles } from './orphanCleanupService';
import { prepareSoftDeletePurgePlan } from './softDeletePurgeService';
import type { IntegrityIssue, IntegrityReport, PendingRestoreOffer } from '../types/offline';

let cachedReport: IntegrityReport | null = null;
let inFlight: Promise<IntegrityReport> | null = null;

async function collectIssues(): Promise<IntegrityIssue[]> {
  const issues: IntegrityIssue[] = [];

  try {
    await getDatabase();
  } catch (error) {
    issues.push({
      code: 'db_unreadable',
      message: error instanceof Error ? error.message : 'Datenbank nicht lesbar',
      severity: 'error'
    });
    return issues;
  }

  for (const table of TABLES) {
    if (!(await tableExists(table))) {
      issues.push({
        code: 'missing_table',
        message: `Tabelle fehlt: ${table}`,
        severity: 'error'
      });
    }
  }

  try {
    const db = await getDatabase();
    await db.getFirstAsync('SELECT 1 as ok');
  } catch (error) {
    issues.push({
      code: 'db_query_failed',
      message: error instanceof Error ? error.message : 'Testabfrage fehlgeschlagen',
      severity: 'error'
    });
  }

  const layout = await ensureStorageLayout();
  for (const [label, path] of Object.entries(layout)) {
    if (!(await fileExists(path))) {
      issues.push({
        code: 'missing_storage_dir',
        message: `Speicherordner fehlt: ${label}`,
        severity: 'warning'
      });
    }
  }

  if (await tableExists('photos')) {
    const db = await getDatabase();
    const photos = await db.getAllAsync<{ id: string; file_path: string; local_path?: string }>(
      `SELECT id, file_path, local_path FROM photos WHERE deleted_at IS NULL`
    );
    for (const photo of photos) {
      const relative = photo.local_path || photo.file_path;
      const uri = resolveDocumentUri(relative);
      if (!(await fileExists(uri))) {
        issues.push({
          code: 'orphan_photo_meta',
          message: `Fotodatei fehlt für ${photo.id}`,
          severity: 'warning'
        });
      }
    }
  }

  return issues;
}

export async function runStartupIntegrityCheck(): Promise<IntegrityReport> {
  if (cachedReport) return cachedReport;
  if (inFlight) return inFlight;

  inFlight = (async () => {
    await ensureStorageLayout();
    let migrationFailed = false;
    let migrationErrorMessage = '';

    try {
      await runMigrations();
    } catch (error) {
      migrationFailed = true;
      migrationErrorMessage = error instanceof Error ? error.message : 'Migration fehlgeschlagen';
    }

    let issues = migrationFailed
      ? [
          {
            code: 'migration_failed',
            message: migrationErrorMessage,
            severity: 'error' as const
          }
        ]
      : await collectIssues();

    const hasFatal = issues.some((issue) => issue.severity === 'error');
    let pendingRestore: PendingRestoreOffer | null = null;

    if (hasFatal) {
      const latest = await getLatestBackupInfo();
      if (latest) {
        pendingRestore = {
          backupUri: latest.uri,
          backupDate: latest.createdAtIso,
          backupName: latest.name
        };
        issues.push({
          code: 'restore_available',
          message: `Backup vom ${new Date(latest.createdAtIso).toLocaleString('de-DE')} zur Wiederherstellung verfügbar.`,
          severity: 'info'
        });
      }
    }

    const orphans = migrationFailed ? [] : await findOrphanFiles().catch(() => []);
    if (orphans.length > 0) {
      issues.push({
        code: 'orphan_files',
        message: `${orphans.length} Datei(en) ohne aktiven Datenbankeintrag gefunden (nur Meldung).`,
        severity: 'warning'
      });
    }

    if (!migrationFailed) {
      await prepareSoftDeletePurgePlan().catch(() => null);
    }

    const report: IntegrityReport = {
      ok: !hasFatal,
      restoredFromBackup: false,
      pendingRestore,
      issues,
      orphanFiles: orphans.map((item) => item.uri)
    };
    cachedReport = report;
    return report;
  })().finally(() => {
    inFlight = null;
  });

  return inFlight;
}

export async function confirmPendingRestore(backupUri: string): Promise<IntegrityReport> {
  await restoreDatabaseFromBackup(backupUri);
  cachedReport = null;
  await runMigrations();
  const issues = await collectIssues();
  const orphans = await findOrphanFiles().catch(() => []);
  const report: IntegrityReport = {
    ok: !issues.some((issue) => issue.severity === 'error'),
    restoredFromBackup: true,
    pendingRestore: null,
    issues,
    orphanFiles: orphans.map((item) => item.uri)
  };
  cachedReport = report;
  return report;
}

export function declinePendingRestore(): void {
  if (!cachedReport) return;
  cachedReport = {
    ...cachedReport,
    pendingRestore: null
  };
}
