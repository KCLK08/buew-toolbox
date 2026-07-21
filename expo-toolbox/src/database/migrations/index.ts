export type Migration = {
  version: number;
  name: string;
  up: string[];
};

export const migrations: Migration[] = [
  {
    version: 1,
    name: 'initial_offline_schema',
    up: [
      `CREATE TABLE IF NOT EXISTS schema_migrations (
        version INTEGER PRIMARY KEY NOT NULL,
        name TEXT NOT NULL,
        applied_at TEXT NOT NULL
      );`,
      `CREATE TABLE IF NOT EXISTS app_meta (
        key TEXT PRIMARY KEY NOT NULL,
        value TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );`,
      `CREATE TABLE IF NOT EXISTS projects (
        id TEXT PRIMARY KEY NOT NULL,
        name TEXT NOT NULL,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        deleted_at TEXT
      );`,
      `CREATE TABLE IF NOT EXISTS diary_runs (
        id TEXT PRIMARY KEY NOT NULL,
        project_id TEXT,
        title TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'draft',
        payload_json TEXT NOT NULL DEFAULT '{}',
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        deleted_at TEXT,
        FOREIGN KEY (project_id) REFERENCES projects(id)
      );`,
      `CREATE TABLE IF NOT EXISTS defects (
        id TEXT PRIMARY KEY NOT NULL,
        project_id TEXT,
        diary_run_id TEXT,
        title TEXT NOT NULL,
        notes TEXT NOT NULL DEFAULT '',
        status TEXT NOT NULL DEFAULT 'draft',
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        deleted_at TEXT,
        FOREIGN KEY (project_id) REFERENCES projects(id),
        FOREIGN KEY (diary_run_id) REFERENCES diary_runs(id)
      );`,
      `CREATE TABLE IF NOT EXISTS notes (
        id TEXT PRIMARY KEY NOT NULL,
        project_id TEXT,
        diary_run_id TEXT,
        defect_id TEXT,
        body TEXT NOT NULL DEFAULT '',
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        deleted_at TEXT,
        FOREIGN KEY (project_id) REFERENCES projects(id),
        FOREIGN KEY (diary_run_id) REFERENCES diary_runs(id),
        FOREIGN KEY (defect_id) REFERENCES defects(id)
      );`,
      `CREATE TABLE IF NOT EXISTS photos (
        id TEXT PRIMARY KEY NOT NULL,
        project_id TEXT,
        diary_run_id TEXT,
        defect_id TEXT,
        file_path TEXT NOT NULL,
        mime_type TEXT NOT NULL,
        byte_size INTEGER NOT NULL DEFAULT 0,
        status TEXT NOT NULL DEFAULT 'ready',
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        deleted_at TEXT,
        FOREIGN KEY (project_id) REFERENCES projects(id),
        FOREIGN KEY (diary_run_id) REFERENCES diary_runs(id),
        FOREIGN KEY (defect_id) REFERENCES defects(id)
      );`,
      `CREATE INDEX IF NOT EXISTS idx_diary_runs_project ON diary_runs(project_id);`,
      `CREATE INDEX IF NOT EXISTS idx_defects_project ON defects(project_id);`,
      `CREATE INDEX IF NOT EXISTS idx_photos_diary_run ON photos(diary_run_id);`,
      `CREATE INDEX IF NOT EXISTS idx_photos_deleted ON photos(deleted_at);`,
      `CREATE INDEX IF NOT EXISTS idx_diary_runs_deleted ON diary_runs(deleted_at);`
    ]
  },
  {
    version: 2,
    name: 'domain_parity_v2',
    up: [
      // Projects — Paritätsfelder
      `ALTER TABLE projects ADD COLUMN description TEXT NOT NULL DEFAULT '';`,
      `ALTER TABLE projects ADD COLUMN location TEXT NOT NULL DEFAULT '';`,
      `ALTER TABLE projects ADD COLUMN date TEXT;`,
      `ALTER TABLE projects ADD COLUMN status TEXT NOT NULL DEFAULT 'active';`,

      // Diary — entry_date (Wetter/Personal/… bleiben in payload_json)
      `ALTER TABLE diary_runs ADD COLUMN entry_date TEXT;`,

      // Defects — description + priority (notes → description kopieren)
      `ALTER TABLE defects ADD COLUMN description TEXT NOT NULL DEFAULT '';`,
      `ALTER TABLE defects ADD COLUMN priority TEXT NOT NULL DEFAULT 'normal';`,
      `ALTER TABLE defects ADD COLUMN diary_entry_id TEXT;`,
      `UPDATE defects SET description = notes WHERE description = '' AND notes IS NOT NULL AND notes != '';`,
      `UPDATE defects SET diary_entry_id = diary_run_id WHERE diary_entry_id IS NULL;`,

      // Notes — diary_entry_id Alias
      `ALTER TABLE notes ADD COLUMN diary_entry_id TEXT;`,
      `UPDATE notes SET diary_entry_id = diary_run_id WHERE diary_entry_id IS NULL;`,

      // Photos — einheitliche Parent-/Datei-Felder
      `ALTER TABLE photos ADD COLUMN parent_id TEXT;`,
      `ALTER TABLE photos ADD COLUMN parent_type TEXT;`,
      `ALTER TABLE photos ADD COLUMN filename TEXT NOT NULL DEFAULT '';`,
      `ALTER TABLE photos ADD COLUMN local_path TEXT NOT NULL DEFAULT '';`,
      `ALTER TABLE photos ADD COLUMN file_size INTEGER NOT NULL DEFAULT 0;`,
      `ALTER TABLE photos ADD COLUMN diary_entry_id TEXT;`,
      `UPDATE photos SET local_path = file_path WHERE (local_path IS NULL OR local_path = '') AND file_path IS NOT NULL;`,
      `UPDATE photos SET file_size = byte_size WHERE file_size = 0;`,
      `UPDATE photos SET diary_entry_id = diary_run_id WHERE diary_entry_id IS NULL;`,
      `UPDATE photos SET parent_id = COALESCE(defect_id, diary_run_id, project_id) WHERE parent_id IS NULL;`,
      `UPDATE photos SET parent_type = CASE
          WHEN defect_id IS NOT NULL THEN 'defect'
          WHEN diary_run_id IS NOT NULL THEN 'diary_entry'
          WHEN project_id IS NOT NULL THEN 'project'
          ELSE parent_type
        END
        WHERE parent_type IS NULL;`,
      `UPDATE photos SET filename = CASE
          WHEN filename != '' THEN filename
          WHEN instr(file_path, '/') > 0 THEN substr(file_path, instr(file_path, '/') + 1)
          ELSE COALESCE(file_path, id)
        END;`,

      // Documents
      `CREATE TABLE IF NOT EXISTS documents (
        id TEXT PRIMARY KEY NOT NULL,
        parent_id TEXT,
        parent_type TEXT,
        filename TEXT NOT NULL,
        local_path TEXT NOT NULL,
        mime_type TEXT NOT NULL,
        file_size INTEGER NOT NULL DEFAULT 0,
        status TEXT NOT NULL DEFAULT 'ready',
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        deleted_at TEXT
      );`,
      `CREATE INDEX IF NOT EXISTS idx_photos_parent ON photos(parent_type, parent_id);`,
      `CREATE INDEX IF NOT EXISTS idx_documents_parent ON documents(parent_type, parent_id);`,
      `CREATE INDEX IF NOT EXISTS idx_documents_deleted ON documents(deleted_at);`
    ]
  }
];
