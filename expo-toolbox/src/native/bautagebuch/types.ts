export type BautagebuchRunStatus = 'draft' | 'completed' | 'deleted';

export type PhotoDocEntry = {
  id: string;
  createdAt: string;
  mimeType: string;
  localPath?: string;
};

export type PhotoDocMeta = {
  enabled: boolean | null;
  entries: PhotoDocEntry[];
  updatedAt: string;
};

export type BautagebuchRun = {
  runId: string;
  templateId: string;
  title: string;
  setupVersion: number;
  values: Record<string, unknown>;
  sectionIndex: number;
  status: BautagebuchRunStatus;
  photoDoc: PhotoDocMeta;
  createdAt: string;
  updatedAt: string;
  completedAt: string;
  deleted_at: string | null;
};

export type BautagebuchExport = {
  exportId: string;
  runId: string;
  fileName: string;
  filePath: string;
  exportedAt: string;
  deleted_at: string | null;
};

export type BautagebuchTemplate = {
  templateId: string;
  templateName: string;
  fileName: string;
  templateKind: string;
  mimeType: string;
  sizeBytes: number;
  pageCount: number;
  pdfPath: string;
  status: 'draft' | 'ready';
  createdAt: string;
  updatedAt: string;
  deleted_at: string | null;
};

export type DetectedField = {
  id: string;
  templateId: string;
  fieldId: string;
  fieldName: string;
  labelCandidate: string;
  type: string;
  options: string[];
  page: number;
  orderIndex: number;
  rect: number[] | null;
  createdAt: string;
  updatedAt: string;
};

export type SetupModelRecord = {
  templateId: string;
  status: 'draft' | 'ready';
  version: number;
  setupModel: Record<string, unknown>;
  createdAt: string;
  updatedAt: string;
};
