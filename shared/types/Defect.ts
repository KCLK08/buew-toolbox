import type { DefectPriority, EntityStatus, SoftDeletable, Timestamps } from './common';

/**
 * Defect / Mangel.
 */
export type Defect = SoftDeletable &
  Timestamps & {
    id: string;
    project_id: string | null;
    diary_entry_id: string | null;
    title: string;
    description: string;
    priority: DefectPriority;
    status: EntityStatus;
  };

export type DefectCreateInput = {
  id?: string;
  project_id?: string | null;
  diary_entry_id?: string | null;
  title: string;
  description?: string;
  priority?: DefectPriority;
  status?: EntityStatus;
};

export type DefectUpdateInput = {
  id: string;
  project_id?: string | null;
  diary_entry_id?: string | null;
  title?: string;
  description?: string;
  priority?: DefectPriority;
  status?: EntityStatus;
};
