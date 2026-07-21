import type { EntityStatus, SoftDeletable, Timestamps } from './common';

/**
 * Structured diary payload (JSON). Same keys on Expo and PWA.
 */
export type DiaryPayload = {
  weather?: {
    category?: string;
    temp_min?: string;
    temp_max?: string;
    notes?: string;
  };
  personnel?: Array<{ name?: string; role?: string; hours?: string }>;
  equipment?: Array<{ name?: string; count?: string; notes?: string }>;
  work?: Array<{ description?: string; location?: string; notes?: string }>;
  notes?: string;
  [key: string]: unknown;
};

/**
 * Diary entry / Bautagebuch-Lauf.
 * Physical table Expo: `diary_runs` (legacy name).
 * Physical store PWA domain layer: `diary_entries`.
 */
export type DiaryEntry = SoftDeletable &
  Timestamps & {
    id: string;
    project_id: string | null;
    title: string;
    /** ISO date (YYYY-MM-DD) or null */
    entry_date: string | null;
    status: EntityStatus;
    /** Serialized DiaryPayload as JSON string for SQLite; object allowed in IndexedDB adapters */
    payload_json: string;
  };

export type DiaryEntryCreateInput = {
  id?: string;
  project_id?: string | null;
  title: string;
  entry_date?: string | null;
  status?: EntityStatus;
  payload_json?: string;
};

export type DiaryEntryUpdateInput = {
  id: string;
  project_id?: string | null;
  title?: string;
  entry_date?: string | null;
  status?: EntityStatus;
  payload_json?: string;
};
