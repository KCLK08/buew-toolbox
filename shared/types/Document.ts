import type { DocumentStatus, ParentType, SoftDeletable, Timestamps } from './common';

/**
 * Document metadata (PDFs, exports references, etc.).
 * Binary stays in platform storage, not in the shared row shape.
 */
export type Document = SoftDeletable &
  Timestamps & {
    id: string;
    parent_id: string | null;
    parent_type: ParentType | null;
    filename: string;
    local_path: string;
    mime_type: string;
    file_size: number;
    status: DocumentStatus;
  };

export type DocumentCreateInput = {
  id?: string;
  parent_id?: string | null;
  parent_type?: ParentType | null;
  filename: string;
  local_path: string;
  mime_type: string;
  file_size?: number;
  status?: DocumentStatus;
  created_at?: string;
};
