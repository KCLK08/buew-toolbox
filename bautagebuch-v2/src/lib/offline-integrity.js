/**
 * Startup integrity checks for Bautagebuch IndexedDB.
 */
export async function runBautagebuchIntegrityCheck(db) {
  const issues = [];

  try {
    await db.open();
  } catch (error) {
    return {
      ok: false,
      issues: [
        {
          code: 'db_unreadable',
          message: error?.message || 'IndexedDB nicht lesbar',
          severity: 'error'
        }
      ]
    };
  }

  const requiredTables = ['templates', 'runs', 'photo_assets', 'exports', 'db_backups'];
  for (const name of requiredTables) {
    if (!db.tables.some((table) => table.name === name)) {
      issues.push({
        code: 'missing_table',
        message: `Tabelle fehlt: ${name}`,
        severity: 'error'
      });
    }
  }

  const domainTables = ['projects', 'diary_entries', 'defects', 'notes', 'photos', 'documents', 'app_meta'];
  for (const name of domainTables) {
    if (!db.tables.some((table) => table.name === name)) {
      issues.push({
        code: 'missing_domain_table',
        message: `Domain-Tabelle fehlt: ${name}`,
        severity: 'warning'
      });
    }
  }

  try {
    const runs = await db.runs.toArray();
    const assets = await db.photo_assets.toArray();
    const activeRuns = new Set(
      runs.filter((run) => !run?.deleted_at).map((run) => String(run.runId || '').trim()).filter(Boolean)
    );

    for (const asset of assets) {
      if (asset?.deleted_at) continue;
      const runId = String(asset.runId || '').trim();
      if (runId && !activeRuns.has(runId)) {
        issues.push({
          code: 'orphan_photo',
          message: `Foto ohne aktiven Lauf: ${asset.id || asset.entryId || 'unbekannt'}`,
          severity: 'warning'
        });
      }
      if (!asset?.data && !asset?.deleted_at) {
        issues.push({
          code: 'photo_missing_bytes',
          message: `Fotodaten fehlen: ${asset.id || asset.entryId || 'unbekannt'}`,
          severity: 'warning'
        });
      }
    }
  } catch (error) {
    issues.push({
      code: 'integrity_scan_failed',
      message: error?.message || 'Integritätsprüfung fehlgeschlagen',
      severity: 'error'
    });
  }

  return {
    ok: !issues.some((issue) => issue.severity === 'error'),
    issues
  };
}
