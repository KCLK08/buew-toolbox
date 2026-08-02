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
  /** Custom display names chosen during step-2 field assignment. */
  fieldLabels?: Record<string, string>;
  /** True after the user has seen the step-1 introduction screen. */
  structureIntroSeen?: boolean;
  /** True after the user has seen the step-2 introduction screen. */
  assignIntroSeen?: boolean;
  /** True after the user has seen the step-3 introduction screen. */
  fieldsIntroSeen?: boolean;
  /** Field settings targets reviewed in step 3 (target keys). */
  configuredFieldIds?: string[];
  /** Active index in the step-3 field settings walkthrough. */
  currentFieldSettingsIndex?: number;
  /** True after the user has completed the full setup wizard once. */
  setupCompleted?: boolean;
  /** True while editing a completed template (skips onboarding screens). */
  editMode?: boolean;
};

export type SetupFieldType =
  | 'text'
  | 'number'
  | 'datetime'
  | 'checkbox'
  | 'select'
  | 'static_text'
  | 'signature'
  | 'table'
  | 'weather';

export type SetupWeatherMetric =
  | 'temperature'
  | 'temperature_min'
  | 'temperature_max'
  | 'condition'
  | 'cloud_cover'
  | 'precipitation'
  | 'humidity'
  | 'wind_direction'
  | 'wind_speed';

export type SetupFieldDateMode = 'date' | 'time' | 'datetime';

export type SetupSignatureMode = 'draw' | 'image';

export type FieldSettingsTarget =
  | {
      kind: 'single';
      sectionId: string;
      fieldId: string;
      key: string;
    }
  | {
      kind: 'table-cell';
      tableId: string;
      rowId: string;
      columnId: string;
      fieldId: string;
      key: string;
    }
  | {
      kind: 'table-meta';
      tableId: string;
      key: string;
    };

export type FieldSource = 'acroform' | 'manual' | 'ocr';

export type FieldRect = {
  x: number;
  y: number;
  width: number;
  height: number;
};

export type FieldGeometry = {
  page: number;
  rect: FieldRect;
};

export type TemplateField = {
  id: string;
  fieldId: string;
  name: string;
  type: SetupFieldType | string;
  geometry: FieldGeometry | null;
  source: FieldSource;
  groupId?: string;
  tableId?: string;
  options: string[];
  orderIndex: number;
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
  geometry?: FieldGeometry | null;
  source?: FieldSource;
  placeholder?: string;
  useCurrentDate?: boolean;
  dateMode?: SetupFieldDateMode;
  checkboxExclusiveGroup?: string;
  signatureMode?: SetupSignatureMode;
  signatureImageUri?: string;
  staticText?: string;
  weatherMetric?: SetupWeatherMetric;
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
  /** Legacy bounding box [x1,y1,x2,y2] in PDF points — derived from geometry when present. */
  rect: number[] | null;
  geometry: FieldGeometry | null;
  source: FieldSource;
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
