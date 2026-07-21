import type { SQLiteDatabase } from 'expo-sqlite';

import { getDatabase, withTransaction } from '../database/sqlite';
import { createUuid, nowIso } from '../lib/ids';
import { createDatabaseBackup } from '../storage/backupService';
import type {
  Defect,
  DefectCreateInput,
  DefectUpdateInput,
  DiaryEntry,
  DiaryEntryCreateInput,
  DiaryEntryUpdateInput,
  Document,
  DocumentCreateInput,
  Note,
  NoteCreateInput,
  NoteUpdateInput,
  ParentType,
  Photo,
  PhotoCreateInput,
  PhotoFilter,
  Project,
  ProjectCreateInput,
  ProjectUpdateInput
} from '../types/offline';

async function importantWrite<T>(work: (db: SQLiteDatabase) => Promise<T>, backup = true): Promise<T> {
  const result = await withTransaction(work);
  if (backup) {
    await createDatabaseBackup().catch(() => null);
  }
  return result;
}

function filenameFromPath(path: string, fallback: string): string {
  const normalized = String(path || '');
  const idx = normalized.lastIndexOf('/');
  return (idx >= 0 ? normalized.slice(idx + 1) : normalized) || fallback;
}

function deriveParent(
  projectId: string | null,
  diaryEntryId: string | null,
  defectId: string | null
): { parent_id: string | null; parent_type: ParentType | null } {
  if (defectId) return { parent_id: defectId, parent_type: 'defect' };
  if (diaryEntryId) return { parent_id: diaryEntryId, parent_type: 'diary_entry' };
  if (projectId) return { parent_id: projectId, parent_type: 'project' };
  return { parent_id: null, parent_type: null };
}

type ProjectRow = Project & { name: string };
type DiaryRow = DiaryEntry & { diary_run_id?: string };
type DefectRow = Defect & {
  diary_run_id?: string | null;
  notes?: string;
};
type NoteRow = Note & { diary_run_id?: string | null };
type PhotoRow = Photo & {
  project_id?: string | null;
  diary_run_id?: string | null;
  diary_entry_id?: string | null;
  defect_id?: string | null;
  file_path?: string;
  byte_size?: number;
};

function mapProject(row: ProjectRow): Project {
  return {
    id: row.id,
    name: row.name,
    description: row.description ?? '',
    location: row.location ?? '',
    date: row.date ?? null,
    status: row.status ?? 'active',
    created_at: row.created_at,
    updated_at: row.updated_at,
    deleted_at: row.deleted_at ?? null
  };
}

function mapDiary(row: DiaryRow): DiaryEntry {
  return {
    id: row.id,
    project_id: row.project_id ?? null,
    title: row.title,
    entry_date: row.entry_date ?? null,
    status: row.status ?? 'draft',
    payload_json: row.payload_json ?? '{}',
    created_at: row.created_at,
    updated_at: row.updated_at,
    deleted_at: row.deleted_at ?? null
  };
}

function mapDefect(row: DefectRow): Defect {
  return {
    id: row.id,
    project_id: row.project_id ?? null,
    diary_entry_id: row.diary_entry_id ?? row.diary_run_id ?? null,
    title: row.title,
    description: row.description || row.notes || '',
    priority: row.priority ?? 'normal',
    status: row.status ?? 'draft',
    created_at: row.created_at,
    updated_at: row.updated_at,
    deleted_at: row.deleted_at ?? null
  };
}

function mapNote(row: NoteRow): Note {
  return {
    id: row.id,
    project_id: row.project_id ?? null,
    diary_entry_id: row.diary_entry_id ?? row.diary_run_id ?? null,
    defect_id: row.defect_id ?? null,
    body: row.body ?? '',
    created_at: row.created_at,
    updated_at: row.updated_at,
    deleted_at: row.deleted_at ?? null
  };
}

function mapPhoto(row: PhotoRow): Photo {
  const localPath = row.local_path || row.file_path || '';
  return {
    id: row.id,
    parent_id: row.parent_id ?? null,
    parent_type: (row.parent_type as ParentType | null) ?? null,
    filename: row.filename || filenameFromPath(localPath, row.id),
    local_path: localPath,
    mime_type: row.mime_type,
    file_size: Number(row.file_size ?? row.byte_size ?? 0),
    status: row.status ?? 'ready',
    created_at: row.created_at,
    updated_at: row.updated_at,
    deleted_at: row.deleted_at ?? null
  };
}

function mapDocument(row: Document): Document {
  return {
    id: row.id,
    parent_id: row.parent_id ?? null,
    parent_type: row.parent_type ?? null,
    filename: row.filename,
    local_path: row.local_path,
    mime_type: row.mime_type,
    file_size: Number(row.file_size ?? 0),
    status: row.status ?? 'ready',
    created_at: row.created_at,
    updated_at: row.updated_at,
    deleted_at: row.deleted_at ?? null
  };
}

export const projectRepository = {
  async getProjects(): Promise<Project[]> {
    const db = await getDatabase();
    const rows = await db.getAllAsync<ProjectRow>(
      `SELECT * FROM projects WHERE deleted_at IS NULL ORDER BY updated_at DESC`
    );
    return rows.map(mapProject);
  },

  async getProjectById(id: string): Promise<Project | null> {
    const db = await getDatabase();
    const row = await db.getFirstAsync<ProjectRow>(
      `SELECT * FROM projects WHERE id = ? AND deleted_at IS NULL`,
      id
    );
    return row ? mapProject(row) : null;
  },

  async createProject(input: ProjectCreateInput): Promise<Project> {
    return this.updateProject({
      id: input.id ?? (await createUuid()),
      name: input.name,
      description: input.description,
      location: input.location,
      date: input.date,
      status: input.status ?? 'active'
    });
  },

  async updateProject(input: ProjectUpdateInput): Promise<Project> {
    return importantWrite(async (db) => {
      const timestamp = nowIso();
      const existing = await db.getFirstAsync<ProjectRow>(`SELECT * FROM projects WHERE id = ?`, input.id);
      const record: Project = {
        id: input.id,
        name: (input.name ?? existing?.name ?? 'Projekt').trim() || 'Projekt',
        description: input.description ?? existing?.description ?? '',
        location: input.location ?? existing?.location ?? '',
        date: input.date !== undefined ? input.date : (existing?.date ?? null),
        status: input.status ?? existing?.status ?? 'active',
        created_at: existing?.created_at ?? timestamp,
        updated_at: timestamp,
        deleted_at: null
      };
      await db.runAsync(
        `INSERT INTO projects (id, name, description, location, date, status, created_at, updated_at, deleted_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL)
         ON CONFLICT(id) DO UPDATE SET
           name = excluded.name,
           description = excluded.description,
           location = excluded.location,
           date = excluded.date,
           status = excluded.status,
           updated_at = excluded.updated_at,
           deleted_at = NULL`,
        record.id,
        record.name,
        record.description,
        record.location,
        record.date,
        record.status,
        record.created_at,
        record.updated_at
      );
      return record;
    });
  },

  async softDeleteProject(id: string): Promise<void> {
    await importantWrite(async (db) => {
      await db.runAsync(`UPDATE projects SET deleted_at = ?, updated_at = ? WHERE id = ?`, nowIso(), nowIso(), id);
    });
  },

  /** @deprecated Use getProjects */
  listActive: () => projectRepository.getProjects(),
  /** @deprecated Use createProject/updateProject */
  upsert: (input: { id?: string; name: string }) =>
    input.id
      ? projectRepository.updateProject({ id: input.id, name: input.name })
      : projectRepository.createProject(input),
  /** @deprecated Use softDeleteProject */
  softDelete: (id: string) => projectRepository.softDeleteProject(id)
};

export const diaryRepository = {
  async getDiaryEntries(projectId?: string): Promise<DiaryEntry[]> {
    const db = await getDatabase();
    const rows = projectId
      ? await db.getAllAsync<DiaryRow>(
          `SELECT * FROM diary_runs WHERE deleted_at IS NULL AND project_id = ? ORDER BY updated_at DESC`,
          projectId
        )
      : await db.getAllAsync<DiaryRow>(
          `SELECT * FROM diary_runs WHERE deleted_at IS NULL ORDER BY updated_at DESC`
        );
    return rows.map(mapDiary);
  },

  async getDiaryEntryById(id: string): Promise<DiaryEntry | null> {
    const db = await getDatabase();
    const row = await db.getFirstAsync<DiaryRow>(
      `SELECT * FROM diary_runs WHERE id = ? AND deleted_at IS NULL`,
      id
    );
    return row ? mapDiary(row) : null;
  },

  async createDiaryEntry(input: DiaryEntryCreateInput): Promise<DiaryEntry> {
    return this.updateDiaryEntry({
      id: input.id ?? (await createUuid()),
      project_id: input.project_id,
      title: input.title,
      entry_date: input.entry_date,
      status: input.status ?? 'draft',
      payload_json: input.payload_json ?? '{}'
    });
  },

  async updateDiaryEntry(input: DiaryEntryUpdateInput): Promise<DiaryEntry> {
    return importantWrite(async (db) => {
      const timestamp = nowIso();
      const existing = await db.getFirstAsync<DiaryRow>(`SELECT * FROM diary_runs WHERE id = ?`, input.id);
      const record: DiaryEntry = {
        id: input.id,
        project_id: input.project_id !== undefined ? input.project_id : (existing?.project_id ?? null),
        title: (input.title ?? existing?.title ?? 'Bautagebuch').trim() || 'Bautagebuch',
        entry_date: input.entry_date !== undefined ? input.entry_date : (existing?.entry_date ?? null),
        status: input.status ?? existing?.status ?? 'draft',
        payload_json: input.payload_json ?? existing?.payload_json ?? '{}',
        created_at: existing?.created_at ?? timestamp,
        updated_at: timestamp,
        deleted_at: null
      };
      await db.runAsync(
        `INSERT INTO diary_runs (id, project_id, title, status, payload_json, entry_date, created_at, updated_at, deleted_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL)
         ON CONFLICT(id) DO UPDATE SET
           project_id = excluded.project_id,
           title = excluded.title,
           status = excluded.status,
           payload_json = excluded.payload_json,
           entry_date = excluded.entry_date,
           updated_at = excluded.updated_at,
           deleted_at = NULL`,
        record.id,
        record.project_id,
        record.title,
        record.status,
        record.payload_json,
        record.entry_date,
        record.created_at,
        record.updated_at
      );
      return record;
    });
  },

  async softDeleteDiaryEntry(id: string): Promise<void> {
    await importantWrite(async (db) => {
      const timestamp = nowIso();
      await db.runAsync(`UPDATE diary_runs SET deleted_at = ?, updated_at = ? WHERE id = ?`, timestamp, timestamp, id);
      await db.runAsync(
        `UPDATE photos SET deleted_at = ?, updated_at = ?, status = 'deleted'
         WHERE deleted_at IS NULL AND (diary_run_id = ? OR diary_entry_id = ? OR (parent_type = 'diary_entry' AND parent_id = ?))`,
        timestamp,
        timestamp,
        id,
        id,
        id
      );
      await db.runAsync(
        `UPDATE defects SET deleted_at = ?, updated_at = ?
         WHERE deleted_at IS NULL AND (diary_run_id = ? OR diary_entry_id = ?)`,
        timestamp,
        timestamp,
        id,
        id
      );
      await db.runAsync(
        `UPDATE notes SET deleted_at = ?, updated_at = ?
         WHERE deleted_at IS NULL AND (diary_run_id = ? OR diary_entry_id = ?)`,
        timestamp,
        timestamp,
        id,
        id
      );
    });
  }
};

/** @deprecated Alias — use diaryRepository */
export const diaryRunRepository = {
  listActive: (projectId?: string) => diaryRepository.getDiaryEntries(projectId),
  get: (id: string) => diaryRepository.getDiaryEntryById(id),
  save: (input: {
    id?: string;
    project_id?: string | null;
    title: string;
    status?: DiaryEntry['status'];
    payload_json?: string;
  }) =>
    input.id
      ? diaryRepository.updateDiaryEntry(input as DiaryEntryUpdateInput)
      : diaryRepository.createDiaryEntry(input),
  softDelete: (id: string) => diaryRepository.softDeleteDiaryEntry(id)
};

export const defectRepository = {
  async getDefects(projectId?: string): Promise<Defect[]> {
    const db = await getDatabase();
    const rows = projectId
      ? await db.getAllAsync<DefectRow>(
          `SELECT * FROM defects WHERE deleted_at IS NULL AND project_id = ? ORDER BY updated_at DESC`,
          projectId
        )
      : await db.getAllAsync<DefectRow>(
          `SELECT * FROM defects WHERE deleted_at IS NULL ORDER BY updated_at DESC`
        );
    return rows.map(mapDefect);
  },

  async getDefectById(id: string): Promise<Defect | null> {
    const db = await getDatabase();
    const row = await db.getFirstAsync<DefectRow>(
      `SELECT * FROM defects WHERE id = ? AND deleted_at IS NULL`,
      id
    );
    return row ? mapDefect(row) : null;
  },

  async createDefect(input: DefectCreateInput): Promise<Defect> {
    return this.updateDefect({
      id: input.id ?? (await createUuid()),
      ...input
    });
  },

  async updateDefect(input: DefectUpdateInput): Promise<Defect> {
    return importantWrite(async (db) => {
      const timestamp = nowIso();
      const existing = await db.getFirstAsync<DefectRow>(`SELECT * FROM defects WHERE id = ?`, input.id);
      const diaryEntryId =
        input.diary_entry_id !== undefined
          ? input.diary_entry_id
          : (existing?.diary_entry_id ?? existing?.diary_run_id ?? null);
      const record: Defect = {
        id: input.id,
        project_id: input.project_id !== undefined ? input.project_id : (existing?.project_id ?? null),
        diary_entry_id: diaryEntryId,
        title: (input.title ?? existing?.title ?? 'Mangel').trim() || 'Mangel',
        description: input.description ?? existing?.description ?? existing?.notes ?? '',
        priority: input.priority ?? existing?.priority ?? 'normal',
        status: input.status ?? existing?.status ?? 'draft',
        created_at: existing?.created_at ?? timestamp,
        updated_at: timestamp,
        deleted_at: null
      };
      await db.runAsync(
        `INSERT INTO defects (
           id, project_id, diary_run_id, diary_entry_id, title, notes, description, priority, status, created_at, updated_at, deleted_at
         ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL)
         ON CONFLICT(id) DO UPDATE SET
           project_id = excluded.project_id,
           diary_run_id = excluded.diary_run_id,
           diary_entry_id = excluded.diary_entry_id,
           title = excluded.title,
           notes = excluded.notes,
           description = excluded.description,
           priority = excluded.priority,
           status = excluded.status,
           updated_at = excluded.updated_at,
           deleted_at = NULL`,
        record.id,
        record.project_id,
        record.diary_entry_id,
        record.diary_entry_id,
        record.title,
        record.description,
        record.description,
        record.priority,
        record.status,
        record.created_at,
        record.updated_at
      );
      return record;
    });
  },

  async softDeleteDefect(id: string): Promise<void> {
    await importantWrite(async (db) => {
      const timestamp = nowIso();
      await db.runAsync(`UPDATE defects SET deleted_at = ?, updated_at = ? WHERE id = ?`, timestamp, timestamp, id);
      await db.runAsync(
        `UPDATE photos SET deleted_at = ?, updated_at = ?, status = 'deleted'
         WHERE deleted_at IS NULL AND (defect_id = ? OR (parent_type = 'defect' AND parent_id = ?))`,
        timestamp,
        timestamp,
        id,
        id
      );
    });
  },

  /** @deprecated */
  save: (input: {
    id?: string;
    project_id?: string | null;
    diary_run_id?: string | null;
    title: string;
    notes?: string;
    status?: Defect['status'];
  }) =>
    input.id
      ? defectRepository.updateDefect({
          id: input.id,
          project_id: input.project_id,
          diary_entry_id: input.diary_run_id,
          title: input.title,
          description: input.notes,
          status: input.status
        })
      : defectRepository.createDefect({
          project_id: input.project_id,
          diary_entry_id: input.diary_run_id,
          title: input.title,
          description: input.notes,
          status: input.status
        }),
  softDelete: (id: string) => defectRepository.softDeleteDefect(id)
};

export const noteRepository = {
  async getNotes(): Promise<Note[]> {
    const db = await getDatabase();
    const rows = await db.getAllAsync<NoteRow>(
      `SELECT * FROM notes WHERE deleted_at IS NULL ORDER BY updated_at DESC`
    );
    return rows.map(mapNote);
  },

  async createNote(input: NoteCreateInput): Promise<Note> {
    return this.updateNote({
      id: input.id ?? (await createUuid()),
      ...input
    });
  },

  async updateNote(input: NoteUpdateInput): Promise<Note> {
    return importantWrite(async (db) => {
      const timestamp = nowIso();
      const existing = await db.getFirstAsync<NoteRow>(`SELECT * FROM notes WHERE id = ?`, input.id);
      const diaryEntryId =
        input.diary_entry_id !== undefined
          ? input.diary_entry_id
          : (existing?.diary_entry_id ?? existing?.diary_run_id ?? null);
      const record: Note = {
        id: input.id,
        project_id: input.project_id !== undefined ? input.project_id : (existing?.project_id ?? null),
        diary_entry_id: diaryEntryId,
        defect_id: input.defect_id !== undefined ? input.defect_id : (existing?.defect_id ?? null),
        body: input.body ?? existing?.body ?? '',
        created_at: existing?.created_at ?? timestamp,
        updated_at: timestamp,
        deleted_at: null
      };
      await db.runAsync(
        `INSERT INTO notes (
           id, project_id, diary_run_id, diary_entry_id, defect_id, body, created_at, updated_at, deleted_at
         ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL)
         ON CONFLICT(id) DO UPDATE SET
           project_id = excluded.project_id,
           diary_run_id = excluded.diary_run_id,
           diary_entry_id = excluded.diary_entry_id,
           defect_id = excluded.defect_id,
           body = excluded.body,
           updated_at = excluded.updated_at,
           deleted_at = NULL`,
        record.id,
        record.project_id,
        record.diary_entry_id,
        record.diary_entry_id,
        record.defect_id,
        record.body,
        record.created_at,
        record.updated_at
      );
      return record;
    });
  },

  async softDeleteNote(id: string): Promise<void> {
    await importantWrite(async (db) => {
      await db.runAsync(`UPDATE notes SET deleted_at = ?, updated_at = ? WHERE id = ?`, nowIso(), nowIso(), id);
    });
  },

  /** @deprecated */
  save: (input: {
    id?: string;
    project_id?: string | null;
    diary_run_id?: string | null;
    defect_id?: string | null;
    body: string;
  }) =>
    input.id
      ? noteRepository.updateNote({
          id: input.id,
          project_id: input.project_id,
          diary_entry_id: input.diary_run_id,
          defect_id: input.defect_id,
          body: input.body
        })
      : noteRepository.createNote({
          project_id: input.project_id,
          diary_entry_id: input.diary_run_id,
          defect_id: input.defect_id,
          body: input.body
        })
};

export const photoRepository = {
  async getPhotos(filter: PhotoFilter = {}): Promise<Photo[]> {
    const db = await getDatabase();
    if (filter.parent_id && filter.parent_type) {
      const rows = await db.getAllAsync<PhotoRow>(
        `SELECT * FROM photos
         WHERE deleted_at IS NULL
           AND (
             (parent_id = ? AND parent_type = ?)
             OR (? = 'diary_entry' AND (diary_run_id = ? OR diary_entry_id = ?))
             OR (? = 'defect' AND defect_id = ?)
             OR (? = 'project' AND project_id = ?)
           )
         ORDER BY created_at DESC`,
        filter.parent_id,
        filter.parent_type,
        filter.parent_type,
        filter.parent_id,
        filter.parent_id,
        filter.parent_type,
        filter.parent_id,
        filter.parent_type,
        filter.parent_id
      );
      return rows.map(mapPhoto);
    }
    const rows = await db.getAllAsync<PhotoRow>(
      `SELECT * FROM photos WHERE deleted_at IS NULL ORDER BY created_at DESC`
    );
    return rows.map(mapPhoto);
  },

  async addPhoto(input: PhotoCreateInput): Promise<Photo> {
    return importantWrite(async (db) => {
      const timestamp = nowIso();
      const id = input.id ?? (await createUuid());
      const parent = {
        parent_id: input.parent_id ?? null,
        parent_type: input.parent_type ?? null
      };
      const projectId = parent.parent_type === 'project' ? parent.parent_id : null;
      const diaryEntryId = parent.parent_type === 'diary_entry' ? parent.parent_id : null;
      const defectId = parent.parent_type === 'defect' ? parent.parent_id : null;
      const record: Photo = {
        id,
        parent_id: parent.parent_id,
        parent_type: parent.parent_type,
        filename: input.filename || filenameFromPath(input.local_path, id),
        local_path: input.local_path,
        mime_type: input.mime_type,
        file_size: Number(input.file_size ?? 0),
        status: input.status ?? 'ready',
        created_at: input.created_at ?? timestamp,
        updated_at: timestamp,
        deleted_at: null
      };
      await db.runAsync(
        `INSERT INTO photos (
           id, project_id, diary_run_id, diary_entry_id, defect_id,
           parent_id, parent_type, filename, local_path, file_path,
           mime_type, byte_size, file_size, status, created_at, updated_at, deleted_at
         ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL)
         ON CONFLICT(id) DO UPDATE SET
           project_id = excluded.project_id,
           diary_run_id = excluded.diary_run_id,
           diary_entry_id = excluded.diary_entry_id,
           defect_id = excluded.defect_id,
           parent_id = excluded.parent_id,
           parent_type = excluded.parent_type,
           filename = excluded.filename,
           local_path = excluded.local_path,
           file_path = excluded.file_path,
           mime_type = excluded.mime_type,
           byte_size = excluded.byte_size,
           file_size = excluded.file_size,
           status = excluded.status,
           updated_at = excluded.updated_at,
           deleted_at = NULL`,
        record.id,
        projectId,
        diaryEntryId,
        diaryEntryId,
        defectId,
        record.parent_id,
        record.parent_type,
        record.filename,
        record.local_path,
        record.local_path,
        record.mime_type,
        record.file_size,
        record.file_size,
        record.status,
        record.created_at,
        record.updated_at
      );
      return record;
    });
  },

  async deletePhoto(id: string): Promise<void> {
    await importantWrite(async (db) => {
      await db.runAsync(
        `UPDATE photos SET deleted_at = ?, updated_at = ?, status = 'deleted' WHERE id = ?`,
        nowIso(),
        nowIso(),
        id
      );
    });
  },

  async listActivePaths(): Promise<string[]> {
    const db = await getDatabase();
    const rows = await db.getAllAsync<{ local_path: string; file_path: string }>(
      `SELECT local_path, file_path FROM photos WHERE deleted_at IS NULL`
    );
    return rows.map((row) => row.local_path || row.file_path).filter(Boolean);
  },

  /** @deprecated */
  listForDiaryRun: (diaryRunId: string) =>
    photoRepository.getPhotos({ parent_id: diaryRunId, parent_type: 'diary_entry' }),
  /** @deprecated */
  saveMetadata: async (
    input: Omit<Photo, 'deleted_at' | 'updated_at' | 'parent_id' | 'parent_type' | 'filename' | 'local_path' | 'file_size'> & {
      project_id?: string | null;
      diary_run_id?: string | null;
      defect_id?: string | null;
      file_path: string;
      byte_size: number;
      updated_at?: string;
    }
  ) => {
    const parent = deriveParent(input.project_id ?? null, input.diary_run_id ?? null, input.defect_id ?? null);
    return photoRepository.addPhoto({
      id: input.id,
      parent_id: parent.parent_id,
      parent_type: parent.parent_type,
      filename: filenameFromPath(input.file_path, input.id),
      local_path: input.file_path,
      mime_type: input.mime_type,
      file_size: input.byte_size,
      status: input.status,
      created_at: input.created_at
    });
  },
  softDelete: (id: string) => photoRepository.deletePhoto(id)
};

export const documentRepository = {
  async getDocuments(filter: PhotoFilter = {}): Promise<Document[]> {
    const db = await getDatabase();
    if (filter.parent_id && filter.parent_type) {
      const rows = await db.getAllAsync<Document>(
        `SELECT * FROM documents WHERE deleted_at IS NULL AND parent_id = ? AND parent_type = ? ORDER BY created_at DESC`,
        filter.parent_id,
        filter.parent_type
      );
      return rows.map(mapDocument);
    }
    const rows = await db.getAllAsync<Document>(
      `SELECT * FROM documents WHERE deleted_at IS NULL ORDER BY created_at DESC`
    );
    return rows.map(mapDocument);
  },

  async addDocument(input: DocumentCreateInput): Promise<Document> {
    return importantWrite(async (db) => {
      const timestamp = nowIso();
      const id = input.id ?? (await createUuid());
      const record: Document = {
        id,
        parent_id: input.parent_id ?? null,
        parent_type: input.parent_type ?? null,
        filename: input.filename,
        local_path: input.local_path,
        mime_type: input.mime_type,
        file_size: Number(input.file_size ?? 0),
        status: input.status ?? 'ready',
        created_at: input.created_at ?? timestamp,
        updated_at: timestamp,
        deleted_at: null
      };
      await db.runAsync(
        `INSERT INTO documents (
           id, parent_id, parent_type, filename, local_path, mime_type, file_size, status, created_at, updated_at, deleted_at
         ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL)
         ON CONFLICT(id) DO UPDATE SET
           parent_id = excluded.parent_id,
           parent_type = excluded.parent_type,
           filename = excluded.filename,
           local_path = excluded.local_path,
           mime_type = excluded.mime_type,
           file_size = excluded.file_size,
           status = excluded.status,
           updated_at = excluded.updated_at,
           deleted_at = NULL`,
        record.id,
        record.parent_id,
        record.parent_type,
        record.filename,
        record.local_path,
        record.mime_type,
        record.file_size,
        record.status,
        record.created_at,
        record.updated_at
      );
      return record;
    });
  },

  async softDeleteDocument(id: string): Promise<void> {
    await importantWrite(async (db) => {
      await db.runAsync(
        `UPDATE documents SET deleted_at = ?, updated_at = ?, status = 'deleted' WHERE id = ?`,
        nowIso(),
        nowIso(),
        id
      );
    });
  }
};
