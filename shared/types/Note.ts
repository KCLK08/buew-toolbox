import type { SoftDeletable, Timestamps } from './common';

/**
 * Free-form note linked to project / diary / defect.
 */
export type Note = SoftDeletable &
  Timestamps & {
    id: string;
    project_id: string | null;
    diary_entry_id: string | null;
    defect_id: string | null;
    body: string;
  };

export type NoteCreateInput = {
  id?: string;
  project_id?: string | null;
  diary_entry_id?: string | null;
  defect_id?: string | null;
  body: string;
};

export type NoteUpdateInput = {
  id: string;
  project_id?: string | null;
  diary_entry_id?: string | null;
  defect_id?: string | null;
  body?: string;
};
