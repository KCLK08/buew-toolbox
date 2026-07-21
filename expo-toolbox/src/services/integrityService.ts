import { TABLES } from '../database/schema/constants';
import { getDatabase, runMigrations, tableExists } from '../database/sqlite';
import { restoreDatabaseFromLatestBackup } from '../storage/backupService';
import { ensureStorageLayout, fileExists, resolveDocumentUri } from '../storage/fileService';
import type { IntegrityIssue, IntegrityReport } from '../types/offline';

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
      message: error instanceof Error ? error.message : 'Tesabfrage fehlgeschlagen',
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
    let restoredFromBackup = false;

    try {
      await runMigrations();
    } catch (error) {
      const restored = await restoreDatabaseFromLatestBackup();
      restoredFromBackup = restored;
      if (!restored) {
        const report: IntegrityReport = {
          ok: false,
          restoredFromBackup: false,
          issues: [
            {
              code: 'migration_failed',
              message: error instanceof Error ? error.message : 'Migration fehlgeschlagen',
              severity: 'error'
            }
          ]
        };
        cachedReport = report;
        return report;
      }
      await runMigrations();
    }

    let issues = await collectIssues();
    const hasFatal = issues.some((issue) => issue.severity === 'error');

    if (hasFatal && !restoredFromBackup) {
      const restored = await restoreDatabaseFromLatestBackup();
      if (restored) {
        restoredFromBackup = true;
        await runMigrations();
        issues = await collectIssues();
      }
    }

    const report: IntegrityReport = {
      ok: !issues.some((issue) => issue.severity === 'error'),
      restoredFromBackup,
      issues
    };
    cachedReport = report;
    return report;
  })().finally(() => {
    inFlight = null;
  });

  return inFlight;
}
