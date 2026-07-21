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
  }
];
