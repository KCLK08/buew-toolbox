/**
 * Expo offline types — re-export shared domain types.
 * Platform-specific integrity types stay local.
 */
export type {
  Defect,
  DefectCreateInput,
  DefectPriority,
  DefectUpdateInput,
  DiaryEntry,
  DiaryEntryCreateInput,
  DiaryEntryUpdateInput,
  DiaryPayload,
  Document,
  DocumentCreateInput,
  DocumentStatus,
  EntityStatus,
  Note,
  NoteCreateInput,
  NoteUpdateInput,
  ParentType,
  Photo,
  PhotoCreateInput,
  PhotoFilter,
  PhotoStatus,
  Project,
  ProjectCreateInput,
  ProjectUpdateInput,
  SoftDeletable,
  Timestamps
} from '@buew/shared/types';

export { DOMAIN_SCHEMA_VERSION as SHARED_DOMAIN_SCHEMA_VERSION } from '@buew/shared/types';

/** @deprecated Use Project */
export type ProjectRecord = import('@buew/shared/types').Project;
/** @deprecated Use DiaryEntry */
export type DiaryRunRecord = import('@buew/shared/types').DiaryEntry;
/** @deprecated Use Defect — maps diary_run_id legacy via repository */
export type DefectRecord = import('@buew/shared/types').Defect;
/** @deprecated Use Note */
export type NoteRecord = import('@buew/shared/types').Note;
/** @deprecated Use Photo */
export type PhotoRecord = import('@buew/shared/types').Photo;

export type IntegrityIssue = {
  code: string;
  message: string;
  severity: 'info' | 'warning' | 'error';
};

export type PendingRestoreOffer = {
  backupUri: string;
  backupDate: string;
  backupName: string;
};

export type IntegrityReport = {
  ok: boolean;
  restoredFromBackup: boolean;
  pendingRestore: PendingRestoreOffer | null;
  issues: IntegrityIssue[];
  orphanFiles: string[];
};
