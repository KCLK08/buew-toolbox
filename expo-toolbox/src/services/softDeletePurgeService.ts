import { getDatabase } from '../database/sqlite';
import { nowIso } from '../lib/ids';
import type { DefectRecord, DiaryRunRecord, NoteRecord, PhotoRecord, ProjectRecord } from '../types/offline';

export type SoftDeletePurgeCandidate = {
  table: 'projects' | 'diary_runs' | 'defects' | 'notes' | 'photos';
  id: string;
  deleted_at: string;
};

export type SoftDeletePurgePlan = {
  dryRun: true;
  olderThanIso: string;
  candidates: SoftDeletePurgeCandidate[];
  note: string;
};

/**
 * Prepare a soft-delete purge plan.
 * Does NOT delete anything — purge remains disabled until explicitly activated later.
 */
export async function prepareSoftDeletePurgePlan(options?: {
  olderThanDays?: number;
}): Promise<SoftDeletePurgePlan> {
  const olderThanDays = options?.olderThanDays ?? 30;
  const cutoff = new Date(Date.now() - olderThanDays * 24 * 60 * 60 * 1000).toISOString();
  const db = await getDatabase();

  const [projects, runs, defects, notes, photos] = await Promise.all([
    db.getAllAsync<ProjectRecord>(
      `SELECT * FROM projects WHERE deleted_at IS NOT NULL AND deleted_at <= ?`,
      cutoff
    ),
    db.getAllAsync<DiaryRunRecord>(
      `SELECT * FROM diary_runs WHERE deleted_at IS NOT NULL AND deleted_at <= ?`,
      cutoff
    ),
    db.getAllAsync<DefectRecord>(
      `SELECT * FROM defects WHERE deleted_at IS NOT NULL AND deleted_at <= ?`,
      cutoff
    ),
    db.getAllAsync<NoteRecord>(
      `SELECT * FROM notes WHERE deleted_at IS NOT NULL AND deleted_at <= ?`,
      cutoff
    ),
    db.getAllAsync<PhotoRecord>(
      `SELECT * FROM photos WHERE deleted_at IS NOT NULL AND deleted_at <= ?`,
      cutoff
    )
  ]);

  const candidates: SoftDeletePurgeCandidate[] = [
    ...projects.map((row) => ({
      table: 'projects' as const,
      id: row.id,
      deleted_at: String(row.deleted_at)
    })),
    ...runs.map((row) => ({
      table: 'diary_runs' as const,
      id: row.id,
      deleted_at: String(row.deleted_at)
    })),
    ...defects.map((row) => ({
      table: 'defects' as const,
      id: row.id,
      deleted_at: String(row.deleted_at)
    })),
    ...notes.map((row) => ({
      table: 'notes' as const,
      id: row.id,
      deleted_at: String(row.deleted_at)
    })),
    ...photos.map((row) => ({
      table: 'photos' as const,
      id: row.id,
      deleted_at: String(row.deleted_at)
    }))
  ];

  return {
    dryRun: true,
    olderThanIso: cutoff,
    candidates,
    note: `Purge ist deaktiviert. Plan erstellt am ${nowIso()} für Einträge älter als ${olderThanDays} Tage.`
  };
}

/** Placeholder — hard purge is intentionally not implemented yet. */
export async function executeSoftDeletePurge(): Promise<never> {
  throw new Error('Soft-Delete-Purge ist noch nicht aktiviert.');
}
