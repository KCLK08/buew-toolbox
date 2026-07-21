import type { ParentType, PhotoStatus, SoftDeletable, Timestamps } from './common';

/**
 * Photo metadata. Binary assets live outside the row
 * (Expo: documentDirectory; PWA: IndexedDB photo_assets / blob store).
 */
export type Photo = SoftDeletable &
  Timestamps & {
    id: string;
    parent_id: string | null;
    parent_type: ParentType | null;
    filename: string;
    /** Relative local path (Expo) or logical key (PWA) */
    local_path: string;
    mime_type: string;
    file_size: number;
    status: PhotoStatus;
  };

export type PhotoCreateInput = {
  id?: string;
  parent_id?: string | null;
  parent_type?: ParentType | null;
  filename: string;
  local_path: string;
  mime_type: string;
  file_size?: number;
  status?: PhotoStatus;
  created_at?: string;
};

export type PhotoFilter = {
  parent_id?: string;
  parent_type?: ParentType;
};
