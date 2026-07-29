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

export type BautagebuchTemplateStatus = 'draft' | 'in_progress' | 'ready' | 'archived';

export type BautagebuchTemplate = {
  templateId: string;
  templateName: string;
  fileName: string;
  templateKind: string;
  mimeType: string;
  sizeBytes: number;
  pageCount: number;
  pdfPath: string;
  status: BautagebuchTemplateStatus;
  createdAt: string;
  updatedAt: string;
  deleted_at: string | null;
};

export type SetupWizardStep = 'structure' | 'assign' | 'fields';

export type SetupWizardGroup = {
  sectionId: string;
  label: string;
  description?: string;
};

export type SetupWizardTableColumn = {
  columnId: string;
  label: string;
  type: 'text' | 'checkbox';
  required?: boolean;
  multiline?: boolean;
  skipped?: boolean;
};

export type SetupWizardTable = {
  tableId: string;
  label: string;
  columns: SetupWizardTableColumn[];
  rowCount: number;
};

export type SetupWizardTableAssignment = {
  tableId: string;
  rowIndex: number;
  columnId: string;
};

export type SetupStructureGroup = {
  id: string;
  name: string;
  description?: string;
  type: 'group';
  order: number;
};

export type SetupStructureTableColumn = {
  id: string;
  name: string;
  order: number;
};

export type SetupStructureTable = {
  id: string;
  name: string;
  type: 'table';
  columns: SetupStructureTableColumn[];
  order: number;
};

export type SetupStructureItem = SetupStructureGroup | SetupStructureTable;

export type SetupWizardState = {
  step: SetupWizardStep;
  currentFieldIndex: number;
  structure: SetupStructureItem[];
  groups: SetupWizardGroup[];
  tables: SetupWizardTable[];
  assignments: Record<string, string>;
  tableAssignments: Record<string, SetupWizardTableAssignment>;
  deferredFieldIds: string[];
  /** True after the user has seen the step-1 introduction screen. */
  structureIntroSeen?: boolean;
};

export type SetupFieldConfig = {
  fieldId: string;
  fieldName?: string;
  label?: string;
  required?: boolean;
  skipped?: boolean;
  multiline?: boolean;
  defaultValue?: string;
  hint?: string;
  page?: number;
  type?: string;
  options?: string[];
  rect?: number[] | null;
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
  status: BautagebuchTemplateStatus;
  version: number;
  setupModel: Record<string, unknown>;
  createdAt: string;
  updatedAt: string;
};
