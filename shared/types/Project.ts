import type { EntityStatus, SoftDeletable, Timestamps } from './common';

/**
 * Project — shared fachliche Entität.
 * Pflicht: id, name, created_at, updated_at
 */
export type Project = SoftDeletable &
  Timestamps & {
    id: string;
    name: string;
    description: string;
    location: string;
    /** ISO date (YYYY-MM-DD) or null */
    date: string | null;
    status: EntityStatus;
  };

export type ProjectCreateInput = {
  id?: string;
  name: string;
  description?: string;
  location?: string;
  date?: string | null;
  status?: EntityStatus;
};

export type ProjectUpdateInput = {
  id: string;
  name?: string;
  description?: string;
  location?: string;
  date?: string | null;
  status?: EntityStatus;
};
