/**
 * Domain schema version — must match shared/types DOMAIN_SCHEMA_VERSION (2).
 * Dexie store format version may be higher (v4+) for additive store rolls.
 */
export const DOMAIN_SCHEMA_VERSION = 2;

export const DOMAIN_STORES_V4 = {
  projects: '&id, status, updated_at, deleted_at',
  diary_entries: '&id, project_id, status, entry_date, updated_at, deleted_at',
  defects: '&id, project_id, diary_entry_id, status, priority, updated_at, deleted_at',
  notes: '&id, project_id, diary_entry_id, defect_id, updated_at, deleted_at',
  photos: '&id, parent_id, parent_type, status, updated_at, deleted_at',
  documents: '&id, parent_id, parent_type, status, updated_at, deleted_at',
  app_meta: '&key'
};
