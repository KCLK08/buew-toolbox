import type { SQLiteDatabase } from 'expo-sqlite';

import { getDatabase, withTransaction } from '../database/sqlite';
import { createUuid, nowIso } from '../lib/ids';
import { createDatabaseBackup } from '../storage/backupService';
import type {
  DefectRecord,
  DiaryRunRecord,
  NoteRecord,
  PhotoRecord,
  ProjectRecord
} from '../types/offline';

async function importantWrite<T>(work: (db: SQLiteDatabase) => Promise<T>, backup = true): Promise<T> {
  const result = await withTransaction(work);
  if (backup) {
    await createDatabaseBackup().catch(() => null);
  }
  return result;
}

export const projectRepository = {
  async listActive(): Promise<ProjectRecord[]> {
    const db = await getDatabase();
    return db.getAllAsync<ProjectRecord>(
      `SELECT * FROM projects WHERE deleted_at IS NULL ORDER BY updated_at DESC`
    );
  },

  async upsert(input: { id?: string; name: string }): Promise<ProjectRecord> {
    return importantWrite(async (db) => {
      const id = input.id ?? (await createUuid());
      const timestamp = nowIso();
      const existing = await db.getFirstAsync<ProjectRecord>(`SELECT * FROM projects WHERE id = ?`, id);
      const record: ProjectRecord = {
        id,
        name: input.name.trim() || 'Projekt',
        created_at: existing?.created_at ?? timestamp,
        updated_at: timestamp,
        deleted_at: null
      };
      await db.runAsync(
        `INSERT INTO projects (id, name, created_at, updated_at, deleted_at)
         VALUES (?, ?, ?, ?, NULL)
         ON CONFLICT(id) DO UPDATE SET
           name = excluded.name,
           updated_at = excluded.updated_at,
           deleted_at = NULL`,
        record.id,
        record.name,
        record.created_at,
        record.updated_at
      );
      return record;
    });
  },

  async softDelete(id: string): Promise<void> {
    await importantWrite(async (db) => {
      await db.runAsync(`UPDATE projects SET deleted_at = ?, updated_at = ? WHERE id = ?`, nowIso(), nowIso(), id);
    });
  }
};

export const diaryRunRepository = {
  async listActive(projectId?: string): Promise<DiaryRunRecord[]> {
    const db = await getDatabase();
    if (projectId) {
      return db.getAllAsync<DiaryRunRecord>(
        `SELECT * FROM diary_runs WHERE deleted_at IS NULL AND project_id = ? ORDER BY updated_at DESC`,
        projectId
      );
    }
    return db.getAllAsync<DiaryRunRecord>(
      `SELECT * FROM diary_runs WHERE deleted_at IS NULL ORDER BY updated_at DESC`
    );
  },

  async get(id: string): Promise<DiaryRunRecord | null> {
    const db = await getDatabase();
    return (
      (await db.getFirstAsync<DiaryRunRecord>(
        `SELECT * FROM diary_runs WHERE id = ? AND deleted_at IS NULL`,
        id
      )) ?? null
    );
  },

  async save(input: {
    id?: string;
    project_id?: string | null;
    title: string;
    status?: DiaryRunRecord['status'];
    payload_json?: string;
  }): Promise<DiaryRunRecord> {
    return importantWrite(async (db) => {
      const id = input.id ?? (await createUuid());
      const timestamp = nowIso();
      const existing = await db.getFirstAsync<DiaryRunRecord>(`SELECT * FROM diary_runs WHERE id = ?`, id);
      const record: DiaryRunRecord = {
        id,
        project_id: input.project_id ?? existing?.project_id ?? null,
        title: input.title.trim() || 'Bautagebuch',
        status: input.status ?? existing?.status ?? 'draft',
        payload_json: input.payload_json ?? existing?.payload_json ?? '{}',
        created_at: existing?.created_at ?? timestamp,
        updated_at: timestamp,
        deleted_at: null
      };
      await db.runAsync(
        `INSERT INTO diary_runs (id, project_id, title, status, payload_json, created_at, updated_at, deleted_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, NULL)
         ON CONFLICT(id) DO UPDATE SET
           project_id = excluded.project_id,
           title = excluded.title,
           status = excluded.status,
           payload_json = excluded.payload_json,
           updated_at = excluded.updated_at,
           deleted_at = NULL`,
        record.id,
        record.project_id,
        record.title,
        record.status,
        record.payload_json,
        record.created_at,
        record.updated_at
      );
      return record;
    });
  },

  async softDelete(id: string): Promise<void> {
    await importantWrite(async (db) => {
      const timestamp = nowIso();
      await db.runAsync(`UPDATE diary_runs SET deleted_at = ?, updated_at = ? WHERE id = ?`, timestamp, timestamp, id);
      await db.runAsync(`UPDATE photos SET deleted_at = ?, updated_at = ?, status = 'deleted' WHERE diary_run_id = ? AND deleted_at IS NULL`, timestamp, timestamp, id);
      await db.runAsync(`UPDATE defects SET deleted_at = ?, updated_at = ? WHERE diary_run_id = ? AND deleted_at IS NULL`, timestamp, timestamp, id);
      await db.runAsync(`UPDATE notes SET deleted_at = ?, updated_at = ? WHERE diary_run_id = ? AND deleted_at IS NULL`, timestamp, timestamp, id);
    });
  }
};

export const defectRepository = {
  async save(input: {
    id?: string;
    project_id?: string | null;
    diary_run_id?: string | null;
    title: string;
    notes?: string;
    status?: DefectRecord['status'];
  }): Promise<DefectRecord> {
    return importantWrite(async (db) => {
      const id = input.id ?? (await createUuid());
      const timestamp = nowIso();
      const existing = await db.getFirstAsync<DefectRecord>(`SELECT * FROM defects WHERE id = ?`, id);
      const record: DefectRecord = {
        id,
        project_id: input.project_id ?? existing?.project_id ?? null,
        diary_run_id: input.diary_run_id ?? existing?.diary_run_id ?? null,
        title: input.title.trim() || 'Mangel',
        notes: input.notes ?? existing?.notes ?? '',
        status: input.status ?? existing?.status ?? 'draft',
        created_at: existing?.created_at ?? timestamp,
        updated_at: timestamp,
        deleted_at: null
      };
      await db.runAsync(
        `INSERT INTO defects (id, project_id, diary_run_id, title, notes, status, created_at, updated_at, deleted_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL)
         ON CONFLICT(id) DO UPDATE SET
           project_id = excluded.project_id,
           diary_run_id = excluded.diary_run_id,
           title = excluded.title,
           notes = excluded.notes,
           status = excluded.status,
           updated_at = excluded.updated_at,
           deleted_at = NULL`,
        record.id,
        record.project_id,
        record.diary_run_id,
        record.title,
        record.notes,
        record.status,
        record.created_at,
        record.updated_at
      );
      return record;
    });
  },

  async softDelete(id: string): Promise<void> {
    await importantWrite(async (db) => {
      const timestamp = nowIso();
      await db.runAsync(`UPDATE defects SET deleted_at = ?, updated_at = ? WHERE id = ?`, timestamp, timestamp, id);
      await db.runAsync(`UPDATE photos SET deleted_at = ?, updated_at = ?, status = 'deleted' WHERE defect_id = ? AND deleted_at IS NULL`, timestamp, timestamp, id);
    });
  }
};

export const noteRepository = {
  async save(input: {
    id?: string;
    project_id?: string | null;
    diary_run_id?: string | null;
    defect_id?: string | null;
    body: string;
  }): Promise<NoteRecord> {
    return importantWrite(async (db) => {
      const id = input.id ?? (await createUuid());
      const timestamp = nowIso();
      const existing = await db.getFirstAsync<NoteRecord>(`SELECT * FROM notes WHERE id = ?`, id);
      const record: NoteRecord = {
        id,
        project_id: input.project_id ?? existing?.project_id ?? null,
        diary_run_id: input.diary_run_id ?? existing?.diary_run_id ?? null,
        defect_id: input.defect_id ?? existing?.defect_id ?? null,
        body: input.body,
        created_at: existing?.created_at ?? timestamp,
        updated_at: timestamp,
        deleted_at: null
      };
      await db.runAsync(
        `INSERT INTO notes (id, project_id, diary_run_id, defect_id, body, created_at, updated_at, deleted_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, NULL)
         ON CONFLICT(id) DO UPDATE SET
           project_id = excluded.project_id,
           diary_run_id = excluded.diary_run_id,
           defect_id = excluded.defect_id,
           body = excluded.body,
           updated_at = excluded.updated_at,
           deleted_at = NULL`,
        record.id,
        record.project_id,
        record.diary_run_id,
        record.defect_id,
        record.body,
        record.created_at,
        record.updated_at
      );
      return record;
    });
  }
};

export const photoRepository = {
  async listForDiaryRun(diaryRunId: string): Promise<PhotoRecord[]> {
    const db = await getDatabase();
    return db.getAllAsync<PhotoRecord>(
      `SELECT * FROM photos WHERE diary_run_id = ? AND deleted_at IS NULL ORDER BY created_at DESC`,
      diaryRunId
    );
  },

  async saveMetadata(input: Omit<PhotoRecord, 'deleted_at' | 'updated_at'> & { updated_at?: string }): Promise<PhotoRecord> {
    return importantWrite(async (db) => {
      const timestamp = nowIso();
      const record: PhotoRecord = {
        ...input,
        updated_at: input.updated_at ?? timestamp,
        deleted_at: null
      };
      await db.runAsync(
        `INSERT INTO photos (
           id, project_id, diary_run_id, defect_id, file_path, mime_type, byte_size, status, created_at, updated_at, deleted_at
         ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL)
         ON CONFLICT(id) DO UPDATE SET
           project_id = excluded.project_id,
           diary_run_id = excluded.diary_run_id,
           defect_id = excluded.defect_id,
           file_path = excluded.file_path,
           mime_type = excluded.mime_type,
           byte_size = excluded.byte_size,
           status = excluded.status,
           updated_at = excluded.updated_at,
           deleted_at = NULL`,
        record.id,
        record.project_id,
        record.diary_run_id,
        record.defect_id,
        record.file_path,
        record.mime_type,
        record.byte_size,
        record.status,
        record.created_at,
        record.updated_at
      );
      return record;
    });
  },

  async softDelete(id: string): Promise<void> {
    await importantWrite(async (db) => {
      await db.runAsync(
        `UPDATE photos SET deleted_at = ?, updated_at = ?, status = 'deleted' WHERE id = ?`,
        nowIso(),
        nowIso(),
        id
      );
    });
  }
};
