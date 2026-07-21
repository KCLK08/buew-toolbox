export type EntityStatus = 'active' | 'archived' | 'draft';

export type PhotoStatus = 'ready' | 'pending' | 'error' | 'deleted';

export type SoftDeletable = {
  deleted_at: string | null;
};

export type ProjectRecord = SoftDeletable & {
  id: string;
  name: string;
  created_at: string;
  updated_at: string;
};

export type DiaryRunRecord = SoftDeletable & {
  id: string;
  project_id: string | null;
  title: string;
  status: EntityStatus;
  payload_json: string;
  created_at: string;
  updated_at: string;
};

export type DefectRecord = SoftDeletable & {
  id: string;
  project_id: string | null;
  diary_run_id: string | null;
  title: string;
  notes: string;
  status: EntityStatus;
  created_at: string;
  updated_at: string;
};

export type NoteRecord = SoftDeletable & {
  id: string;
  project_id: string | null;
  diary_run_id: string | null;
  defect_id: string | null;
  body: string;
  created_at: string;
  updated_at: string;
};

export type PhotoRecord = SoftDeletable & {
  id: string;
  project_id: string | null;
  diary_run_id: string | null;
  defect_id: string | null;
  file_path: string;
  mime_type: string;
  byte_size: number;
  status: PhotoStatus;
  created_at: string;
  updated_at: string;
};

export type IntegrityIssue = {
  code: string;
  message: string;
  severity: 'info' | 'warning' | 'error';
};

export type IntegrityReport = {
  ok: boolean;
  restoredFromBackup: boolean;
  issues: IntegrityIssue[];
};
