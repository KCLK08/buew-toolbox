/**
 * Shared-shaped offline repositories for Bautagebuch PWA (IndexedDB).
 * Method names and record shapes match Expo repositories / shared types.
 */

import { DOMAIN_SCHEMA_VERSION } from './domain-schema.js';
import { ensureOfflineDbReady, getOfflineDb } from './db.js';

function nowIso() {
  return new Date().toISOString();
}

function createId() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return `id_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

function isActive(record) {
  return !String(record?.deleted_at || '').trim();
}

async function setDomainSchemaVersion(db) {
  await db.app_meta.put({
    key: 'domain_schema_version',
    value: String(DOMAIN_SCHEMA_VERSION),
    updated_at: nowIso()
  });
}

export const projectRepository = {
  async getProjects() {
    await ensureOfflineDbReady();
    const rows = await getOfflineDb().projects.toArray();
    return rows.filter(isActive).sort((a, b) => String(b.updated_at).localeCompare(String(a.updated_at)));
  },

  async getProjectById(id) {
    await ensureOfflineDbReady();
    const row = await getOfflineDb().projects.get(id);
    return row && isActive(row) ? row : null;
  },

  async createProject(input) {
    return this.updateProject({
      id: input.id || createId(),
      name: input.name,
      description: input.description ?? '',
      location: input.location ?? '',
      date: input.date ?? null,
      status: input.status ?? 'active'
    });
  },

  async updateProject(input) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const existing = await db.projects.get(input.id);
    const timestamp = nowIso();
    const record = {
      id: input.id,
      name: String(input.name ?? existing?.name ?? 'Projekt').trim() || 'Projekt',
      description: input.description ?? existing?.description ?? '',
      location: input.location ?? existing?.location ?? '',
      date: input.date !== undefined ? input.date : (existing?.date ?? null),
      status: input.status ?? existing?.status ?? 'active',
      created_at: existing?.created_at ?? timestamp,
      updated_at: timestamp,
      deleted_at: null
    };
    await db.projects.put(record);
    await setDomainSchemaVersion(db);
    return record;
  },

  async softDeleteProject(id) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const existing = await db.projects.get(id);
    if (!existing) return;
    await db.projects.put({
      ...existing,
      deleted_at: nowIso(),
      updated_at: nowIso()
    });
  }
};

export const diaryRepository = {
  async getDiaryEntries(projectId) {
    await ensureOfflineDbReady();
    let rows = await getOfflineDb().diary_entries.toArray();
    rows = rows.filter(isActive);
    if (projectId) {
      rows = rows.filter((row) => row.project_id === projectId);
    }
    return rows.sort((a, b) => String(b.updated_at).localeCompare(String(a.updated_at)));
  },

  async getDiaryEntryById(id) {
    await ensureOfflineDbReady();
    const row = await getOfflineDb().diary_entries.get(id);
    return row && isActive(row) ? row : null;
  },

  async createDiaryEntry(input) {
    return this.updateDiaryEntry({
      id: input.id || createId(),
      project_id: input.project_id ?? null,
      title: input.title,
      entry_date: input.entry_date ?? null,
      status: input.status ?? 'draft',
      payload_json: input.payload_json ?? '{}'
    });
  },

  async updateDiaryEntry(input) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const existing = await db.diary_entries.get(input.id);
    const timestamp = nowIso();
    const record = {
      id: input.id,
      project_id: input.project_id !== undefined ? input.project_id : (existing?.project_id ?? null),
      title: String(input.title ?? existing?.title ?? 'Bautagebuch').trim() || 'Bautagebuch',
      entry_date: input.entry_date !== undefined ? input.entry_date : (existing?.entry_date ?? null),
      status: input.status ?? existing?.status ?? 'draft',
      payload_json: input.payload_json ?? existing?.payload_json ?? '{}',
      created_at: existing?.created_at ?? timestamp,
      updated_at: timestamp,
      deleted_at: null
    };
    await db.diary_entries.put(record);
    return record;
  },

  async softDeleteDiaryEntry(id) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const timestamp = nowIso();
    const existing = await db.diary_entries.get(id);
    if (existing) {
      await db.diary_entries.put({ ...existing, deleted_at: timestamp, updated_at: timestamp });
    }
    const relatedPhotos = await db.photos
      .filter((row) => row.parent_type === 'diary_entry' && row.parent_id === id && isActive(row))
      .toArray();
    for (const photo of relatedPhotos) {
      await db.photos.put({ ...photo, deleted_at: timestamp, updated_at: timestamp, status: 'deleted' });
    }
    const relatedDefects = await db.defects.filter((row) => row.diary_entry_id === id && isActive(row)).toArray();
    for (const defect of relatedDefects) {
      await db.defects.put({ ...defect, deleted_at: timestamp, updated_at: timestamp });
    }
    const relatedNotes = await db.notes.filter((row) => row.diary_entry_id === id && isActive(row)).toArray();
    for (const note of relatedNotes) {
      await db.notes.put({ ...note, deleted_at: timestamp, updated_at: timestamp });
    }
  }
};

export const defectRepository = {
  async getDefects(projectId) {
    await ensureOfflineDbReady();
    let rows = await getOfflineDb().defects.toArray();
    rows = rows.filter(isActive);
    if (projectId) rows = rows.filter((row) => row.project_id === projectId);
    return rows.sort((a, b) => String(b.updated_at).localeCompare(String(a.updated_at)));
  },

  async getDefectById(id) {
    await ensureOfflineDbReady();
    const row = await getOfflineDb().defects.get(id);
    return row && isActive(row) ? row : null;
  },

  async createDefect(input) {
    return this.updateDefect({
      id: input.id || createId(),
      ...input
    });
  },

  async updateDefect(input) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const existing = await db.defects.get(input.id);
    const timestamp = nowIso();
    const record = {
      id: input.id,
      project_id: input.project_id !== undefined ? input.project_id : (existing?.project_id ?? null),
      diary_entry_id:
        input.diary_entry_id !== undefined ? input.diary_entry_id : (existing?.diary_entry_id ?? null),
      title: String(input.title ?? existing?.title ?? 'Mangel').trim() || 'Mangel',
      description: input.description ?? existing?.description ?? '',
      priority: input.priority ?? existing?.priority ?? 'normal',
      status: input.status ?? existing?.status ?? 'draft',
      created_at: existing?.created_at ?? timestamp,
      updated_at: timestamp,
      deleted_at: null
    };
    await db.defects.put(record);
    return record;
  },

  async softDeleteDefect(id) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const timestamp = nowIso();
    const existing = await db.defects.get(id);
    if (existing) {
      await db.defects.put({ ...existing, deleted_at: timestamp, updated_at: timestamp });
    }
    const relatedPhotos = await db.photos
      .filter((row) => row.parent_type === 'defect' && row.parent_id === id && isActive(row))
      .toArray();
    for (const photo of relatedPhotos) {
      await db.photos.put({ ...photo, deleted_at: timestamp, updated_at: timestamp, status: 'deleted' });
    }
  }
};

export const photoRepository = {
  async getPhotos(filter = {}) {
    await ensureOfflineDbReady();
    let rows = await getOfflineDb().photos.toArray();
    rows = rows.filter(isActive);
    if (filter.parent_id) rows = rows.filter((row) => row.parent_id === filter.parent_id);
    if (filter.parent_type) rows = rows.filter((row) => row.parent_type === filter.parent_type);
    return rows.sort((a, b) => String(b.created_at).localeCompare(String(a.created_at)));
  },

  async addPhoto(input) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const timestamp = nowIso();
    const id = input.id || createId();
    const record = {
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
    await db.photos.put(record);
    return record;
  },

  async deletePhoto(id) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const existing = await db.photos.get(id);
    if (!existing) return;
    await db.photos.put({
      ...existing,
      deleted_at: nowIso(),
      updated_at: nowIso(),
      status: 'deleted'
    });
  }
};

export const documentRepository = {
  async getDocuments(filter = {}) {
    await ensureOfflineDbReady();
    let rows = await getOfflineDb().documents.toArray();
    rows = rows.filter(isActive);
    if (filter.parent_id) rows = rows.filter((row) => row.parent_id === filter.parent_id);
    if (filter.parent_type) rows = rows.filter((row) => row.parent_type === filter.parent_type);
    return rows.sort((a, b) => String(b.created_at).localeCompare(String(a.created_at)));
  },

  async addDocument(input) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const timestamp = nowIso();
    const id = input.id || createId();
    const record = {
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
    await db.documents.put(record);
    return record;
  },

  async softDeleteDocument(id) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const existing = await db.documents.get(id);
    if (!existing) return;
    await db.documents.put({
      ...existing,
      deleted_at: nowIso(),
      updated_at: nowIso(),
      status: 'deleted'
    });
  }
};

export const noteRepository = {
  async getNotes() {
    await ensureOfflineDbReady();
    const rows = await getOfflineDb().notes.toArray();
    return rows.filter(isActive).sort((a, b) => String(b.updated_at).localeCompare(String(a.updated_at)));
  },

  async createNote(input) {
    return this.updateNote({
      id: input.id || createId(),
      ...input
    });
  },

  async updateNote(input) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const existing = await db.notes.get(input.id);
    const timestamp = nowIso();
    const record = {
      id: input.id,
      project_id: input.project_id !== undefined ? input.project_id : (existing?.project_id ?? null),
      diary_entry_id:
        input.diary_entry_id !== undefined ? input.diary_entry_id : (existing?.diary_entry_id ?? null),
      defect_id: input.defect_id !== undefined ? input.defect_id : (existing?.defect_id ?? null),
      body: input.body ?? existing?.body ?? '',
      created_at: existing?.created_at ?? timestamp,
      updated_at: timestamp,
      deleted_at: null
    };
    await db.notes.put(record);
    return record;
  },

  async softDeleteNote(id) {
    await ensureOfflineDbReady();
    const db = getOfflineDb();
    const existing = await db.notes.get(id);
    if (!existing) return;
    await db.notes.put({ ...existing, deleted_at: nowIso(), updated_at: nowIso() });
  }
};
