import type {
  DetectedField,
  FieldSettingsTarget,
  SetupFieldConfig,
  SetupFieldType,
  SetupWizardState
} from '../types';
import { getWizardState, withWizardState } from './setup-mapping';
import { syncSectionOrder } from './setup-model.js';

export type FieldSettingsProgress = {
  total: number;
  configured: number;
  open: number;
  percent: number;
};

export const SETUP_FIELD_TYPE_OPTIONS: Array<{ value: SetupFieldType; label: string }> = [
  { value: 'text', label: 'Textfeld' },
  { value: 'number', label: 'Zahlenfeld' },
  { value: 'datetime', label: 'Datum/Zeitfeld' },
  { value: 'checkbox', label: 'Checkbox' },
  { value: 'select', label: 'Auswahlfeld' },
  { value: 'static_text', label: 'Statischer Text' },
  { value: 'signature', label: 'Unterschrift' },
  { value: 'table', label: 'Tabellenfeld' }
];

const DETECTED_TYPE_LABELS: Record<string, string> = {
  text: 'Textfeld',
  checkbox: 'Checkbox',
  dropdown: 'Auswahlfeld',
  radio: 'Auswahlfeld',
  number: 'Zahlenfeld',
  datetime: 'Datum/Zeitfeld',
  select: 'Auswahlfeld',
  static_text: 'Statischer Text',
  signature: 'Unterschrift',
  table: 'Tabellenfeld'
};

function targetKey(target: FieldSettingsTarget): string {
  return target.key;
}

export function buildFieldSettingsTargetKey(input: {
  kind: FieldSettingsTarget['kind'];
  sectionId?: string;
  tableId?: string;
  fieldId?: string;
}): string {
  if (input.kind === 'single') {
    return `single:${input.sectionId}:${input.fieldId}`;
  }
  if (input.kind === 'table-cell') {
    return `table-cell:${input.tableId}:${input.fieldId}`;
  }
  return `table-meta:${input.tableId}`;
}

export function listFieldSettingsTargets(setupModel: Record<string, unknown>): FieldSettingsTarget[] {
  const singleSections = (Array.isArray(setupModel.single_sections)
    ? setupModel.single_sections
    : []) as Array<{ sectionId?: string; fields?: SetupFieldConfig[] }>;
  const tableSections = (Array.isArray(setupModel.table_sections)
    ? setupModel.table_sections
    : []) as Array<{
    tableId?: string;
    rows?: Array<{
      rowId?: string;
      cells?: Array<{ fieldId?: string; columnId?: string }>;
    }>;
  }>;

  const singleById = new Map(
    singleSections.map((section) => [String(section.sectionId || ''), section])
  );
  const tableById = new Map(tableSections.map((table) => [String(table.tableId || ''), table]));

  const targets: FieldSettingsTarget[] = [];
  for (const entry of syncSectionOrder(setupModel)) {
    if (entry.kind === 'single') {
      const section = singleById.get(entry.id);
      const fields = Array.isArray(section?.fields) ? section.fields : [];
      for (const field of fields) {
        const fieldId = String(field.fieldId || '').trim();
        if (!fieldId) continue;
        targets.push({
          kind: 'single',
          sectionId: entry.id,
          fieldId,
          key: buildFieldSettingsTargetKey({ kind: 'single', sectionId: entry.id, fieldId })
        });
      }
      continue;
    }

    const table = tableById.get(entry.id);
    if (!table) continue;
    const seenFieldIds = new Set<string>();
    for (const row of table.rows || []) {
      for (const cell of row.cells || []) {
        const fieldId = String(cell.fieldId || '').trim();
        if (!fieldId || seenFieldIds.has(fieldId)) continue;
        seenFieldIds.add(fieldId);
        targets.push({
          kind: 'table-cell',
          tableId: entry.id,
          rowId: String(row.rowId || ''),
          columnId: String(cell.columnId || ''),
          fieldId,
          key: buildFieldSettingsTargetKey({ kind: 'table-cell', tableId: entry.id, fieldId })
        });
      }
    }
    if ((table.rows || []).some((row) => (row.cells || []).some((cell) => cell.fieldId))) {
      targets.push({
        kind: 'table-meta',
        tableId: entry.id,
        key: buildFieldSettingsTargetKey({ kind: 'table-meta', tableId: entry.id })
      });
    }
  }

  return targets;
}

export function resolveDetectedFieldTypeLabel(
  field: SetupFieldConfig | null,
  detectedFields: DetectedField[]
): string {
  if (!field) return 'Textfeld';
  const detected = detectedFields.find((entry) => entry.fieldId === field.fieldId);
  const raw = String(detected?.type || field.type || 'text');
  return DETECTED_TYPE_LABELS[raw] || DETECTED_TYPE_LABELS.text;
}

export function normalizeSetupFieldType(
  field: SetupFieldConfig | null,
  detectedFields: DetectedField[] = []
): SetupFieldType {
  const raw = String(field?.type || '').trim();
  if (raw === 'dropdown' || raw === 'radio') return 'select';
  if (raw && SETUP_FIELD_TYPE_OPTIONS.some((option) => option.value === raw)) {
    return raw as SetupFieldType;
  }
  const detected = detectedFields.find((entry) => entry.fieldId === field?.fieldId);
  const detectedType = String(detected?.type || 'text');
  if (detectedType === 'dropdown' || detectedType === 'radio') return 'select';
  if (detectedType === 'checkbox') return 'checkbox';
  return 'text';
}

export function setupFieldTypeLabel(type: SetupFieldType): string {
  return SETUP_FIELD_TYPE_OPTIONS.find((option) => option.value === type)?.label || 'Textfeld';
}

export function getFieldSettingsProgress(
  targets: FieldSettingsTarget[],
  wizard: SetupWizardState
): FieldSettingsProgress {
  const configuredIds = new Set(wizard.configuredFieldIds || []);
  const total = targets.length;
  const configured = targets.filter((target) => configuredIds.has(targetKey(target))).length;
  const open = Math.max(0, total - configured);
  const percent = total > 0 ? Math.round((configured / total) * 100) : 100;
  return { total, configured, open, percent };
}

export function isFieldSettingsTargetConfigured(
  target: FieldSettingsTarget,
  wizard: SetupWizardState
): boolean {
  return (wizard.configuredFieldIds || []).includes(targetKey(target));
}

export function resolveCurrentFieldSettingsIndex(
  targets: FieldSettingsTarget[],
  wizard: SetupWizardState
): number {
  if (targets.length === 0) return 0;
  const stored = Math.max(0, Number(wizard.currentFieldSettingsIndex || 0));
  const storedTarget = targets[stored];
  if (storedTarget && !isFieldSettingsTargetConfigured(storedTarget, wizard)) {
    return stored;
  }
  const nextOpen = targets.findIndex(
    (target) => !isFieldSettingsTargetConfigured(target, wizard)
  );
  return nextOpen >= 0 ? nextOpen : Math.min(stored, targets.length - 1);
}

export function getNextOpenFieldSettingsIndex(
  targets: FieldSettingsTarget[],
  wizard: SetupWizardState,
  startIndex = 0
): number {
  for (let index = Math.max(0, startIndex); index < targets.length; index += 1) {
    if (!isFieldSettingsTargetConfigured(targets[index], wizard)) {
      return index;
    }
  }
  return -1;
}

export function markFieldSettingsTargetConfigured(
  setupModel: Record<string, unknown>,
  target: FieldSettingsTarget
): Record<string, unknown> {
  const wizard = getWizardState(setupModel);
  const key = targetKey(target);
  const configuredFieldIds = wizard.configuredFieldIds || [];
  if (configuredFieldIds.includes(key)) {
    return setupModel;
  }
  return withWizardState(setupModel, {
    configuredFieldIds: [...configuredFieldIds, key]
  });
}

export function resolveFieldFromTarget(
  setupModel: Record<string, unknown>,
  target: FieldSettingsTarget
): SetupFieldConfig | null {
  if (target.kind === 'table-meta') return null;
  if (target.kind === 'single') {
    const sections = Array.isArray(setupModel.single_sections) ? setupModel.single_sections : [];
    for (const section of sections) {
      if (String(section?.sectionId) !== target.sectionId) continue;
      const fields = Array.isArray(section?.fields) ? section.fields : [];
      return (
        (fields.find((field: SetupFieldConfig) => String(field?.fieldId) === target.fieldId) as SetupFieldConfig) ||
        null
      );
    }
    return null;
  }

  const tables = Array.isArray(setupModel.table_sections) ? setupModel.table_sections : [];
  for (const table of tables) {
    if (String(table?.tableId) !== target.tableId) continue;
    for (const row of table?.rows || []) {
      for (const cell of row?.cells || []) {
        if (String(cell?.fieldId) === target.fieldId) {
          return cell as SetupFieldConfig;
        }
      }
    }
  }
  return null;
}

export function resolveTargetGroupLabel(
  setupModel: Record<string, unknown>,
  target: FieldSettingsTarget
): string {
  if (target.kind === 'single') {
    const sections = Array.isArray(setupModel.single_sections) ? setupModel.single_sections : [];
    const section = sections.find((entry) => String(entry?.sectionId) === target.sectionId);
    return String(section?.label || 'Gruppe');
  }
  const tables = Array.isArray(setupModel.table_sections) ? setupModel.table_sections : [];
  const table = tables.find((entry) => String(entry?.tableId) === target.tableId);
  return String(table?.label || 'Tabelle');
}

export function resolveTargetDisplayName(
  setupModel: Record<string, unknown>,
  target: FieldSettingsTarget,
  field: SetupFieldConfig | null
): string {
  if (target.kind === 'table-meta') {
    return resolveTargetGroupLabel(setupModel, target);
  }
  return String(field?.label || field?.fieldName || field?.fieldId || 'Feld').trim();
}

export function updateFieldSettingsTarget(
  setupModel: Record<string, unknown>,
  target: FieldSettingsTarget,
  patch: Partial<SetupFieldConfig>
): Record<string, unknown> {
  if (target.kind === 'single') {
    const singleSections = Array.isArray(setupModel.single_sections)
      ? [...setupModel.single_sections]
      : [];
    const nextSections = singleSections.map((section) => {
      if (String(section?.sectionId) !== target.sectionId) return section;
      const fields = Array.isArray(section?.fields) ? [...section.fields] : [];
      return {
        ...section,
        fields: fields.map((field) =>
          String(field?.fieldId) === target.fieldId ? { ...field, ...patch } : field
        )
      };
    });
    return { ...setupModel, single_sections: nextSections, updatedAt: new Date().toISOString() };
  }

  if (target.kind === 'table-cell') {
    const tableSections = Array.isArray(setupModel.table_sections)
      ? [...setupModel.table_sections]
      : [];
    const nextTables = tableSections.map((table) => {
      if (String(table?.tableId) !== target.tableId) return table;
    const rows = Array.isArray(table?.rows)
      ? table.rows.map((row: { cells?: SetupFieldConfig[] }) => ({
          ...row,
          cells: Array.isArray(row?.cells)
            ? row.cells.map((cell: SetupFieldConfig) =>
                String(cell?.fieldId) === target.fieldId ? { ...cell, ...patch } : cell
              )
            : row?.cells
        }))
      : [];
    const columns = Array.isArray(table?.columns)
      ? table.columns.map((column: SetupFieldConfig & { columnId?: string }) =>
          String(column?.columnId) === target.columnId ? { ...column, ...patch } : column
        )
      : table?.columns;
      return { ...table, rows, columns };
    });
    return { ...setupModel, table_sections: nextTables, updatedAt: new Date().toISOString() };
  }

  return setupModel;
}

export function listCheckboxGroupOptions(
  setupModel: Record<string, unknown>,
  detectedFields: DetectedField[]
): Array<{ id: string; label: string }> {
  const targets = listFieldSettingsTargets(setupModel);
  const options: Array<{ id: string; label: string }> = [];
  for (const target of targets) {
    if (target.kind === 'table-meta') continue;
    const field = resolveFieldFromTarget(setupModel, target);
    if (!field) continue;
    if (normalizeSetupFieldType(field, detectedFields) !== 'checkbox') continue;
    const groupId = String(field.checkboxExclusiveGroup || '').trim();
    if (!groupId) continue;
    if (options.some((entry) => entry.id === groupId)) continue;
    options.push({
      id: groupId,
      label: resolveTargetGroupLabel(setupModel, target)
    });
  }
  return options;
}

export function advanceFieldSettingsWalkthrough(
  setupModel: Record<string, unknown>,
  targets: FieldSettingsTarget[],
  currentTarget: FieldSettingsTarget
): Record<string, unknown> {
  let next = markFieldSettingsTargetConfigured(setupModel, currentTarget);
  const wizard = getWizardState(next);
  const nextIndex = getNextOpenFieldSettingsIndex(targets, wizard, 0);
  return withWizardState(next, {
    currentFieldSettingsIndex: nextIndex >= 0 ? nextIndex : wizard.currentFieldSettingsIndex
  });
}

export function applyFieldTypeChange(
  field: SetupFieldConfig,
  nextType: SetupFieldType,
  detectedFields: DetectedField[]
): Partial<SetupFieldConfig> {
  const detected = detectedFields.find((entry) => entry.fieldId === field.fieldId);
  const patch: Partial<SetupFieldConfig> = { type: nextType };
  if (nextType === 'select' && (!field.options || field.options.length === 0)) {
    patch.options = Array.isArray(detected?.options) ? [...detected.options] : [];
  }
  if (nextType === 'datetime' && !field.dateMode) {
    patch.dateMode = 'date';
  }
  if (nextType === 'signature' && !field.signatureMode) {
    patch.signatureMode = 'draw';
  }
  if (nextType === 'static_text' && !field.staticText) {
    patch.staticText = field.label || field.fieldName || '';
  }
  return patch;
}

export function listTableColumnsForMeta(
  setupModel: Record<string, unknown>,
  tableId: string
): Array<{ columnId: string; label: string; type: string; required?: boolean; skipped?: boolean }> {
  const tables = Array.isArray(setupModel.table_sections) ? setupModel.table_sections : [];
  const table = tables.find((entry) => String(entry?.tableId) === tableId);
  if (!table) return [];
  return (Array.isArray(table.columns) ? table.columns : []).map(
    (column: { columnId?: string; label?: string; type?: string; required?: boolean; skipped?: boolean }) => ({
    columnId: String(column.columnId || ''),
    label: String(column.label || 'Spalte'),
    type: String(column.type || 'text'),
    required: column.required === true,
    skipped: column.skipped === true
  }));
}

export function updateTableColumnMeta(
  setupModel: Record<string, unknown>,
  tableId: string,
  columnId: string,
  patch: Partial<{ label: string; type: string; required: boolean; skipped: boolean }>
): Record<string, unknown> {
  const tableSections = Array.isArray(setupModel.table_sections)
    ? [...setupModel.table_sections]
    : [];
  const nextTables = tableSections.map((table) => {
    if (String(table?.tableId) !== tableId) return table;
    return {
      ...table,
      columns: (Array.isArray(table?.columns) ? table.columns : []).map(
        (column: { columnId?: string; label?: string; type?: string; required?: boolean; skipped?: boolean }) =>
          String(column?.columnId) === columnId ? { ...column, ...patch } : column
      )
    };
  });
  return { ...setupModel, table_sections: nextTables, updatedAt: new Date().toISOString() };
}

export function updateTableMetaLabel(
  setupModel: Record<string, unknown>,
  tableId: string,
  label: string
): Record<string, unknown> {
  const tableSections = Array.isArray(setupModel.table_sections)
    ? [...setupModel.table_sections]
    : [];
  const nextTables = tableSections.map((table) =>
    String(table?.tableId) === tableId ? { ...table, label: label.trim() || table.label } : table
  );
  return { ...setupModel, table_sections: nextTables, updatedAt: new Date().toISOString() };
}

export function updateTableMetaFlags(
  setupModel: Record<string, unknown>,
  tableId: string,
  patch: Partial<{ skipped: boolean; multiline: boolean }>
): Record<string, unknown> {
  const tableSections = Array.isArray(setupModel.table_sections)
    ? [...setupModel.table_sections]
    : [];
  const nextTables = tableSections.map((table) =>
    String(table?.tableId) === tableId ? { ...table, ...patch } : table
  );
  return { ...setupModel, table_sections: nextTables, updatedAt: new Date().toISOString() };
}

export function resolveTableMeta(
  setupModel: Record<string, unknown>,
  tableId: string
): { label: string; skipped?: boolean; multiline?: boolean } | null {
  const tables = Array.isArray(setupModel.table_sections) ? setupModel.table_sections : [];
  const table = tables.find((entry) => String(entry?.tableId) === tableId);
  if (!table) return null;
  return {
    label: String(table.label || 'Tabelle'),
    skipped: table.skipped === true,
    multiline: table.multiline === true
  };
}

export function isFieldSettingsWalkthroughComplete(
  targets: FieldSettingsTarget[],
  wizard: SetupWizardState
): boolean {
  if (targets.length === 0) return true;
  return targets.every((target) => isFieldSettingsTargetConfigured(target, wizard));
}
