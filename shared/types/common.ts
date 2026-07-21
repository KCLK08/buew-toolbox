/**
 * Shared offline domain schema version.
 * Expo SQLite `SCHEMA_VERSION` and PWA Dexie domain migrations must match this value.
 */
export const DOMAIN_SCHEMA_VERSION = 2 as const;

export type EntityStatus = 'draft' | 'active' | 'archived' | 'completed';

export type PhotoStatus = 'ready' | 'pending' | 'error' | 'deleted';

export type DocumentStatus = 'ready' | 'pending' | 'error' | 'deleted';

export type DefectPriority = 'low' | 'normal' | 'high' | 'critical';

/** Polymorphic parent for photos/documents. */
export type ParentType = 'project' | 'diary_entry' | 'defect' | 'note';

export type SoftDeletable = {
  deleted_at: string | null;
};

export type Timestamps = {
  created_at: string;
  updated_at: string;
};
