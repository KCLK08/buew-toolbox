export const DB_NAME = 'buew_toolbox.db';
export const BACKUP_PREFIX = 'buew_toolbox_backup_';
export const MAX_BACKUPS = 3;
export const SCHEMA_VERSION = 1;

export const PHOTO_DIR = 'photos';
export const DOCUMENT_DIR = 'documents';
export const BACKUP_DIR = 'backups';

export const TABLES = [
  'schema_migrations',
  'projects',
  'diary_runs',
  'defects',
  'notes',
  'photos',
  'app_meta'
] as const;
