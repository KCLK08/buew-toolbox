import { DOMAIN_SCHEMA_VERSION } from '@buew/shared/types/common';

export const DB_NAME = 'buew_toolbox.db';
export const SITEREPORT_DB_NAME = 'sitereport_native.db';
export const BAUTAGEBUCH_DB_NAME = 'bautagebuch_v2_native.db';
export const BACKUP_PREFIX = 'buew_toolbox_backup_';
export const SITEREPORT_BACKUP_PREFIX = 'sitereport_native_backup_';
export const BAUTAGEBUCH_BACKUP_PREFIX = 'bautagebuch_v2_native_backup_';
export const MAX_BACKUPS = 3;
/** Must match shared DOMAIN_SCHEMA_VERSION */
export const SCHEMA_VERSION = DOMAIN_SCHEMA_VERSION;

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
  'documents',
  'app_meta'
] as const;
