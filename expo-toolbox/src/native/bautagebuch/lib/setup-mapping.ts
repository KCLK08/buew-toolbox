import type { Href } from 'expo-router';

import type {
  DetectedField,
  FieldGeometry,
  FieldSource,
  SetupFieldConfig,
  SetupStructureItem,
  SetupWizardGroup,
  SetupWizardState,
  SetupWizardStep
} from '../types';
import { updateStructureTable } from './setup-structure';
import { buildTableSectionsFromWizard } from './setup-wizard-tables';
import { getFieldPage, fieldToPreviewLegacyRect } from './template-field';
import { buildLegacySectionOrder, syncSectionOrder } from './setup-model.js';

export type MappingField = {
  fieldId: string;
  fieldName: string;
  labelCandidate: string;
  type: string;
  options: string[];
  page: number;
  orderIndex: number;
  /** Stable 1-based index across all mapping views (independent of geometry). */
  displayOrder: number;
  rect: number[] | null;
  geometry: FieldGeometry | null;
  source: FieldSource;
};

export type MappingProgress = {
  current: number;
  total: number;
  percent: number;
  remaining: number;
  assigned: number;
  open: number;
};

export type OverlayPlacement = 'top' | 'bottom' | 'left' | 'right';

function nowIso(): string {
  return new Date().toISOString();
}

function createId(prefix = 'id'): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

function normalizeStructureOrder(items: SetupStructureItem[]): SetupStructureItem[] {
  return items
    .slice()
    .sort((left, right) => left.order - right.order)
    .map((item, index) => ({ ...item, order: index }));
}

function migrateStructureFromLegacy(
  groups: SetupWizardGroup[],
  tables: SetupWizardState['tables']
): SetupStructureItem[] {
  const items: SetupStructureItem[] = [];
  let order = 0;
  for (const group of groups) {
    items.push({
      id: group.sectionId,
      name: group.label,
      description: group.description,
      type: 'group',
      order: order++
    });
  }
  for (const table of tables) {
    items.push({
      id: table.tableId,
      name: table.label,
      type: 'table',
      order: order++,
      columns: table.columns.map((column, index) => ({
        id: column.columnId,
        name: column.label,
        order: index
      }))
    });
  }
  return items;
}

function parseStructureItems(
  raw: Partial<SetupWizardState>,
  groups: SetupWizardGroup[],
  tables: SetupWizardState['tables']
): SetupStructureItem[] {
  if (Array.isArray(raw.structure) && raw.structure.length > 0) {
    return normalizeStructureOrder(
      raw.structure.map((item) => {
        if (item.type === 'table') {
          return {
            ...item,
            columns: [...(item.columns || [])].sort((left, right) => left.order - right.order)
          };
        }
        return item;
      })
    );
  }
  if (groups.length > 0 || tables.length > 0) {
    return migrateStructureFromLegacy(groups, tables);
  }
  return [];
}

function resolveWizardStep(raw: Partial<SetupWizardState>): SetupWizardStep {
  if (raw.step === 'fields') return 'fields';
  if (raw.step === 'assign') return 'assign';
  if (raw.step === 'structure') return 'structure';
  if (raw.step === 'mapping') {
    const assignmentCount =
      raw.assignments && typeof raw.assignments === 'object'
        ? Object.keys(raw.assignments).length
        : 0;
    const tableAssignmentCount =
      raw.tableAssignments && typeof raw.tableAssignments === 'object'
        ? Object.keys(raw.tableAssignments).length
        : 0;
    const deferredCount = Array.isArray(raw.deferredFieldIds) ? raw.deferredFieldIds.length : 0;
    if (assignmentCount > 0 || tableAssignmentCount > 0 || deferredCount > 0) {
      return 'assign';
    }
    return 'structure';
  }
  return 'structure';
}

export function sortMappingFields(detectedFields: DetectedField[]): MappingField[] {
  return [...detectedFields]
    .filter((field) => String(field.fieldId || '').trim())
    .filter((field) => String(field.type || 'text') !== 'unsupported')
    .sort((left, right) => {
      const pageDelta = Number(left.page || 1) - Number(right.page || 1);
      if (pageDelta !== 0) return pageDelta;
      return Number(left.orderIndex || 0) - Number(right.orderIndex || 0);
    })
    .map((field, index) => ({
      fieldId: String(field.fieldId),
      fieldName: String(field.fieldName || ''),
      labelCandidate: String(field.labelCandidate || field.fieldName || 'Feld'),
      type: String(field.type || 'text'),
      options: Array.isArray(field.options) ? field.options.map((entry) => String(entry)) : [],
      page: getFieldPage(field),
      orderIndex: Number(field.orderIndex || 0),
      displayOrder: index + 1,
      rect: fieldToPreviewLegacyRect(field),
      geometry: field.geometry,
      source: field.source || 'acroform'
    }));
}

export function resolveFieldDisplayLabel(
  field: MappingField,
  wizard?: SetupWizardState,
  draftLabels?: Record<string, string>
): string {
  const draft = draftLabels?.[field.fieldId];
  if (draft !== undefined && draft.trim()) return draft.trim();
  const fromField = String(field.labelCandidate || field.fieldName || '').trim();
  if (fromField) return fromField;
  const custom = wizard?.fieldLabels?.[field.fieldId];
  if (custom && custom.trim()) return custom.trim();
  return 'Feld';
}

/** Label for controlled name inputs — preserves empty drafts while editing. */
export function resolveFieldEditLabel(
  field: MappingField,
  wizard?: SetupWizardState,
  draftLabels?: Record<string, string>
): string {
  const draft = draftLabels?.[field.fieldId];
  if (draft !== undefined) return draft;
  const fromField = String(field.labelCandidate || field.fieldName || '').trim();
  if (fromField) return fromField;
  const custom = wizard?.fieldLabels?.[field.fieldId];
  if (custom && custom.trim()) return custom.trim();
  return '';
}

export function buildFieldPreviewHighlights(
  mappingFields: MappingField[],
  resolveLabel: (field: MappingField) => string
): Array<{
  fieldId: string;
  fieldName: string;
  label: string;
  source: FieldSource;
  index: number;
  page: number;
  rect: number[];
}> {
  return mappingFields
    .filter((field) => field.rect && field.rect.length >= 4)
    .map((field) => ({
      fieldId: field.fieldId,
      fieldName: field.fieldName,
      label: resolveLabel(field),
      source: field.source,
      index: field.displayOrder,
      page: field.page,
      rect: field.rect as number[]
    }));
}

export function getWizardState(setupModel: Record<string, unknown>): SetupWizardState {
  const raw = (setupModel.wizard || {}) as Partial<SetupWizardState>;
  const groups =
    Array.isArray(raw.groups) && raw.groups.length > 0
      ? raw.groups.map((group) => ({
          sectionId: String(group.sectionId || ''),
          label: String(group.label || 'Gruppe'),
          description: group.description ? String(group.description) : undefined
        }))
      : [];
  const tables =
    Array.isArray(raw.tables) && raw.tables.length > 0
      ? raw.tables.map((table) => ({
          tableId: String(table.tableId || ''),
          label: String(table.label || 'Tabelle'),
          rowCount: Math.max(1, Number(table.rowCount || 1)),
          columns: Array.isArray(table.columns)
            ? table.columns.map((column) => ({
                columnId: String(column.columnId || ''),
                label: String(column.label || 'Spalte'),
                type: column.type === 'checkbox' ? ('checkbox' as const) : ('text' as const),
                required: column.required === true,
                multiline: column.multiline === true,
                skipped: column.skipped === true
              }))
            : []
        }))
      : [];

  const structure = parseStructureItems(raw, groups, tables);

  return {
    step: resolveWizardStep(raw),
    currentFieldIndex: Math.max(0, Number(raw.currentFieldIndex || 0)),
    structure,
    groups,
    tables,
    assignments:
      raw.assignments && typeof raw.assignments === 'object'
        ? Object.fromEntries(
            Object.entries(raw.assignments).map(([fieldId, sectionId]) => [
              String(fieldId),
              String(sectionId)
            ])
          )
        : {},
    tableAssignments:
      raw.tableAssignments && typeof raw.tableAssignments === 'object'
        ? Object.fromEntries(
            Object.entries(raw.tableAssignments).map(([fieldId, assignment]) => {
              const entry = assignment as Record<string, unknown>;
              return [
                String(fieldId),
                {
                  tableId: String(entry.tableId || ''),
                  rowIndex: Math.max(0, Number(entry.rowIndex || 0)),
                  columnId: String(entry.columnId || '')
                }
              ];
            })
          )
        : {},
    deferredFieldIds: Array.isArray(raw.deferredFieldIds)
      ? raw.deferredFieldIds.map((fieldId) => String(fieldId))
      : [],
    fieldLabels:
      raw.fieldLabels && typeof raw.fieldLabels === 'object'
        ? Object.fromEntries(
            Object.entries(raw.fieldLabels).map(([fieldId, label]) => [
              String(fieldId),
              String(label)
            ])
          )
        : {},
    structureIntroSeen: raw.structureIntroSeen === true,
    assignIntroSeen: raw.assignIntroSeen === true,
    fieldsIntroSeen: raw.fieldsIntroSeen === true,
    configuredFieldIds: Array.isArray(raw.configuredFieldIds)
      ? raw.configuredFieldIds.map((entry) => String(entry))
      : [],
    currentFieldSettingsIndex: Math.max(0, Number(raw.currentFieldSettingsIndex || 0)),
    setupCompleted: raw.setupCompleted === true,
    editMode: raw.editMode === true
  };
}

export function withWizardState(
  setupModel: Record<string, unknown>,
  wizard: Partial<SetupWizardState>
): Record<string, unknown> {
  const current = getWizardState(setupModel);
  return {
    ...setupModel,
    wizard: {
      ...current,
      ...wizard,
      groups: wizard.groups || current.groups,
      tables: wizard.tables || current.tables,
      structure: wizard.structure || current.structure,
      assignments: wizard.assignments || current.assignments,
      tableAssignments: wizard.tableAssignments || current.tableAssignments,
      deferredFieldIds: wizard.deferredFieldIds || current.deferredFieldIds,
      fieldLabels: wizard.fieldLabels || current.fieldLabels,
      structureIntroSeen:
        wizard.structureIntroSeen !== undefined ? wizard.structureIntroSeen : current.structureIntroSeen,
      assignIntroSeen:
        wizard.assignIntroSeen !== undefined ? wizard.assignIntroSeen : current.assignIntroSeen,
      fieldsIntroSeen:
        wizard.fieldsIntroSeen !== undefined ? wizard.fieldsIntroSeen : current.fieldsIntroSeen,
      configuredFieldIds: wizard.configuredFieldIds || current.configuredFieldIds,
      currentFieldSettingsIndex:
        wizard.currentFieldSettingsIndex !== undefined
          ? wizard.currentFieldSettingsIndex
          : current.currentFieldSettingsIndex,
      setupCompleted:
        wizard.setupCompleted !== undefined ? wizard.setupCompleted : current.setupCompleted,
      editMode: wizard.editMode !== undefined ? wizard.editMode : current.editMode
    },
    updatedAt: nowIso()
  };
}

export function getMappingProgress(
  fields: MappingField[],
  wizard: SetupWizardState
): MappingProgress {
  const total = fields.length;
  const assigned = fields.filter(
    (field) =>
      Boolean(wizard.assignments[field.fieldId]) ||
      Boolean(wizard.tableAssignments[field.fieldId])
  ).length;
  const handled = fields.filter(
    (field) =>
      Boolean(wizard.assignments[field.fieldId]) ||
      Boolean(wizard.tableAssignments[field.fieldId]) ||
      wizard.deferredFieldIds.includes(field.fieldId)
  ).length;
  const open = Math.max(0, total - assigned);
  const current = total > 0 ? Math.min(total, Math.max(1, handled + 1)) : 0;
  const percent = total > 0 ? Math.round((assigned / total) * 100) : 100;
  return {
    current,
    total,
    percent,
    remaining: Math.max(0, total - handled),
    assigned,
    open
  };
}

export function getAssignedFieldIds(wizard: SetupWizardState): string[] {
  const ids = new Set<string>();
  for (const fieldId of Object.keys(wizard.assignments)) ids.add(fieldId);
  for (const fieldId of Object.keys(wizard.tableAssignments)) ids.add(fieldId);
  return [...ids];
}

export function isFieldAssigned(fieldId: string, wizard: SetupWizardState): boolean {
  return (
    Boolean(wizard.assignments[fieldId]) || Boolean(wizard.tableAssignments[fieldId])
  );
}

export type FieldAssignmentSummary = {
  kind: 'group' | 'table' | 'none';
  label: string;
  sectionId?: string;
  tableId?: string;
};

export function resolveFieldAssignmentSummary(
  setupModel: Record<string, unknown>,
  fieldId: string
): FieldAssignmentSummary {
  const wizard = getWizardState(setupModel);
  const groupId = wizard.assignments[fieldId];
  if (groupId) {
    const group = wizard.groups.find((entry) => entry.sectionId === groupId);
    return {
      kind: 'group',
      label: group?.label || 'Gruppe',
      sectionId: groupId
    };
  }
  const tableAssignment = wizard.tableAssignments[fieldId];
  if (tableAssignment) {
    const table = wizard.tables.find((entry) => entry.tableId === tableAssignment.tableId);
    const column = table?.columns.find((entry) => entry.columnId === tableAssignment.columnId);
    const tableLabel = table?.label || 'Tabelle';
    const columnLabel = column?.label ? ` · ${column.label}` : '';
    return {
      kind: 'table',
      label: `${tableLabel}${columnLabel}`,
      tableId: tableAssignment.tableId
    };
  }
  return { kind: 'none', label: '' };
}

function removeFieldFromSingleSections(
  setupModel: Record<string, unknown>,
  fieldId: string
): Record<string, unknown>[] {
  const singleSections = Array.isArray(setupModel.single_sections) ? setupModel.single_sections : [];
  return singleSections.map((section) => {
    const fields = (Array.isArray(section?.fields) ? section.fields : []) as SetupFieldConfig[];
    return {
      ...section,
      fields: fields.filter((field) => String(field.fieldId || '') !== fieldId)
    };
  });
}

function removeFieldFromTableSections(
  setupModel: Record<string, unknown>,
  fieldId: string
): Record<string, unknown>[] {
  const tableSections = Array.isArray(setupModel.table_sections) ? setupModel.table_sections : [];
  return tableSections.map((table) => {
    type TableColumn = {
      columnId?: string;
      label?: string;
      type?: string;
      required?: boolean;
      multiline?: boolean;
    };
    type TableCell = SetupFieldConfig & {
      cellId?: string;
      tableId?: string;
      rowId?: string;
      columnId?: string;
    };
    const columns = (Array.isArray(table?.columns) ? table.columns : []) as TableColumn[];
    const rows = (Array.isArray(table?.rows) ? table.rows : []) as Array<{ cells?: TableCell[] }>;
    return {
      ...table,
      rows: rows.map((row) => {
        const cells = (Array.isArray(row?.cells) ? row.cells : []) as TableCell[];
        return {
          ...row,
          cells: cells.map((cell) => {
            if (String(cell?.fieldId || '') !== fieldId) return cell;
            const column = columns.find(
              (entry) => String(entry?.columnId || '') === String(cell?.columnId || '')
            );
            return {
              ...cell,
              fieldId: '',
              fieldName: '',
              label: String(column?.label || cell?.label || 'Spalte'),
              type: String(column?.type || cell?.type || 'text'),
              options: [],
              page: 1,
              rect: null,
              geometry: null,
              skipped: true,
              required: column?.required === true,
              multiline: column?.multiline === true
            };
          })
        };
      })
    };
  });
}

function removeFieldFromConfiguredTargets(configuredFieldIds: string[], fieldId: string): string[] {
  const suffix = `:${fieldId}`;
  return configuredFieldIds.filter((key) => !String(key).endsWith(suffix));
}

export function removeFieldFromWizard(
  setupModel: Record<string, unknown>,
  fieldId: string,
  fieldCount = 0
): Record<string, unknown> {
  const wizard = getWizardState(setupModel);
  const assignments = { ...wizard.assignments };
  delete assignments[fieldId];
  const tableAssignments = { ...wizard.tableAssignments };
  delete tableAssignments[fieldId];
  const deferredFieldIds = wizard.deferredFieldIds.filter((entry) => entry !== fieldId);
  const fieldLabels = { ...(wizard.fieldLabels || {}) };
  delete fieldLabels[fieldId];
  const configuredFieldIds = removeFieldFromConfiguredTargets(
    wizard.configuredFieldIds || [],
    fieldId
  );
  const maxIndex = Math.max(0, fieldCount - 1);
  const nextIndex = Math.min(wizard.currentFieldIndex, maxIndex);
  const withWizard = withWizardState(setupModel, {
    assignments,
    tableAssignments,
    deferredFieldIds,
    fieldLabels,
    configuredFieldIds,
    currentFieldIndex: nextIndex
  });
  return {
    ...withWizard,
    single_sections: removeFieldFromSingleSections(withWizard, fieldId),
    table_sections: removeFieldFromTableSections(withWizard, fieldId),
    updatedAt: nowIso()
  };
}

export function updateFieldDisplayLabel(
  setupModel: Record<string, unknown>,
  fieldId: string,
  label: string
): Record<string, unknown> {
  const wizard = getWizardState(setupModel);
  const trimmed = String(label || '').trim();
  if (!trimmed) {
    const fieldLabels = { ...(wizard.fieldLabels || {}) };
    delete fieldLabels[fieldId];
    return withWizardState(setupModel, { fieldLabels });
  }
  return withWizardState(setupModel, {
    fieldLabels: { ...wizard.fieldLabels, [fieldId]: trimmed }
  });
}

export function getNextUnassignedIndex(
  fields: MappingField[],
  wizard: SetupWizardState,
  startIndex = 0
): number {
  for (let index = Math.max(0, startIndex); index < fields.length; index += 1) {
    const field = fields[index];
    if (
      !wizard.assignments[field.fieldId] &&
      !wizard.tableAssignments[field.fieldId] &&
      !wizard.deferredFieldIds.includes(field.fieldId)
    ) {
      return index;
    }
  }
  return -1;
}

export function resolveNextSequentialIndex(currentIndex: number, total: number): number {
  if (total <= 0) return 0;
  return Math.min(Math.max(0, currentIndex) + 1, total - 1);
}

export function advanceMappingWalkthrough(
  setupModel: Record<string, unknown>,
  fields: MappingField[],
  currentIndex: number
): Record<string, unknown> {
  const wizard = getWizardState(setupModel);
  return withWizardState(setupModel, {
    currentFieldIndex: resolveNextSequentialIndex(currentIndex, fields.length)
  });
}

export function isMappingComplete(fields: MappingField[], wizard: SetupWizardState): boolean {
  if (fields.length === 0) return false;
  return fields.every(
    (field) =>
      Boolean(wizard.assignments[field.fieldId]) ||
      Boolean(wizard.tableAssignments[field.fieldId]) ||
      wizard.deferredFieldIds.includes(field.fieldId)
  );
}

export function assignFieldToGroup(
  setupModel: Record<string, unknown>,
  fieldId: string,
  sectionId: string,
  label?: string
): Record<string, unknown> {
  const wizard = getWizardState(setupModel);
  const assignments = { ...wizard.assignments, [fieldId]: sectionId };
  const tableAssignments = { ...wizard.tableAssignments };
  delete tableAssignments[fieldId];
  const deferredFieldIds = wizard.deferredFieldIds.filter((entry) => entry !== fieldId);
  const trimmedLabel = String(label || '').trim();
  const fieldLabels =
    trimmedLabel.length > 0
      ? { ...wizard.fieldLabels, [fieldId]: trimmedLabel }
      : wizard.fieldLabels;
  return withWizardState(setupModel, {
    assignments,
    tableAssignments,
    deferredFieldIds,
    fieldLabels
  });
}

export function assignFieldToTableColumn(
  setupModel: Record<string, unknown>,
  fieldId: string,
  tableId: string,
  input: { columnId?: string; newColumnName?: string; fieldLabel?: string }
): Record<string, unknown> {
  let model = setupModel;
  let columnId = input.columnId ? String(input.columnId) : '';

  if (!columnId) {
    const columnName =
      String(input.newColumnName || input.fieldLabel || '').trim() || 'Spalte';
    model = addWizardTableColumn(model, tableId, { label: columnName, type: 'text' });
    const wizardAfterColumn = getWizardState(model);
    const table = wizardAfterColumn.tables.find((entry) => entry.tableId === tableId);
    columnId = table?.columns[table.columns.length - 1]?.columnId || '';
    if (table) {
      model = updateStructureTable(model, tableId, {
        columns: table.columns.map((column) => ({ id: column.columnId, name: column.label }))
      });
    }
  }

  if (!columnId) {
    throw new Error('Spalte konnte nicht erstellt werden.');
  }

  model = assignFieldToTableCell(model, fieldId, {
    tableId,
    rowIndex: 0,
    columnId
  });

  const trimmedLabel = String(input.fieldLabel || '').trim();
  if (trimmedLabel.length > 0) {
    const wizard = getWizardState(model);
    model = withWizardState(model, {
      fieldLabels: { ...wizard.fieldLabels, [fieldId]: trimmedLabel }
    });
  }

  return model;
}

export function deferField(
  setupModel: Record<string, unknown>,
  fieldId: string
): Record<string, unknown> {
  const wizard = getWizardState(setupModel);
  const deferredFieldIds = wizard.deferredFieldIds.includes(fieldId)
    ? wizard.deferredFieldIds
    : [...wizard.deferredFieldIds, fieldId];
  const assignments = { ...wizard.assignments };
  delete assignments[fieldId];
  const tableAssignments = { ...wizard.tableAssignments };
  delete tableAssignments[fieldId];
  return withWizardState(setupModel, { assignments, tableAssignments, deferredFieldIds });
}

export function addWizardTableSection(
  setupModel: Record<string, unknown>,
  label: string
): { setupModel: Record<string, unknown>; table: import('../types').SetupWizardTable } {
  const trimmed = String(label || '').trim() || 'Neue Tabelle';
  const table = {
    tableId: createId('table'),
    label: trimmed,
    columns: [] as import('../types').SetupWizardTableColumn[],
    rowCount: 1
  };
  const wizard = getWizardState(setupModel);
  return {
    setupModel: withWizardState(setupModel, { tables: [...wizard.tables, table] }),
    table
  };
}

export function addWizardTableColumn(
  setupModel: Record<string, unknown>,
  tableId: string,
  input: { label: string; type: 'text' | 'checkbox' }
): Record<string, unknown> {
  const wizard = getWizardState(setupModel);
  const column = {
    columnId: createId('col'),
    label: String(input.label || '').trim() || (input.type === 'checkbox' ? 'Checkbox' : 'Spalte'),
    type: input.type,
    required: false,
    multiline: false,
    skipped: false
  };
  const tables = wizard.tables.map((table) =>
    table.tableId === tableId ? { ...table, columns: [...table.columns, column] } : table
  );
  return withWizardState(setupModel, { tables });
}

export function addWizardTableRow(
  setupModel: Record<string, unknown>,
  tableId: string
): Record<string, unknown> {
  const wizard = getWizardState(setupModel);
  const tables = wizard.tables.map((table) =>
    table.tableId === tableId ? { ...table, rowCount: Math.max(1, table.rowCount + 1) } : table
  );
  return withWizardState(setupModel, { tables });
}

export function assignFieldToTableCell(
  setupModel: Record<string, unknown>,
  fieldId: string,
  assignment: import('../types').SetupWizardTableAssignment
): Record<string, unknown> {
  const wizard = getWizardState(setupModel);
  const tableAssignments = { ...wizard.tableAssignments, [fieldId]: assignment };
  const assignments = { ...wizard.assignments };
  delete assignments[fieldId];
  const deferredFieldIds = wizard.deferredFieldIds.filter((entry) => entry !== fieldId);
  return withWizardState(setupModel, { tableAssignments, assignments, deferredFieldIds });
}

export function addWizardGroup(
  setupModel: Record<string, unknown>,
  label: string
): { setupModel: Record<string, unknown>; group: SetupWizardGroup } {
  const trimmed = String(label || '').trim();
  const group: SetupWizardGroup = {
    sectionId: createId('group'),
    label: trimmed || 'Neue Gruppe'
  };
  const wizard = getWizardState(setupModel);
  return {
    setupModel: withWizardState(setupModel, { groups: [...wizard.groups, group] }),
    group
  };
}

function fieldConfigFromDetected(
  field: MappingField,
  existing?: SetupFieldConfig,
  customLabel?: string
): SetupFieldConfig {
  return {
    fieldId: field.fieldId,
    fieldName: field.fieldName,
    label: String(
      existing?.label || field.labelCandidate || customLabel || field.fieldName || 'Feld'
    ).trim(),
    type: field.type,
    options: Array.isArray(existing?.options)
      ? [...existing.options]
      : Array.isArray(field.options)
        ? [...field.options]
        : [],
    required: Boolean(existing?.required),
    skipped: Boolean(existing?.skipped),
    multiline: Boolean(existing?.multiline),
    defaultValue: String(existing?.defaultValue || ''),
    hint: String(existing?.hint || ''),
    page: field.page,
    rect: field.rect,
    geometry: field.geometry,
    source: field.source
  };
}

export function findExistingFieldConfig(
  setupModel: Record<string, unknown>,
  fieldId: string
): SetupFieldConfig | null {
  const singleSections = Array.isArray(setupModel.single_sections) ? setupModel.single_sections : [];
  for (const section of singleSections) {
    const fields = Array.isArray(section?.fields) ? section.fields : [];
    const match = fields.find((field: SetupFieldConfig) => String(field?.fieldId) === String(fieldId));
    if (match) return match as SetupFieldConfig;
  }
  return null;
}

export function rebuildSectionsFromWizard(
  setupModel: Record<string, unknown>,
  fields: MappingField[]
): Record<string, unknown> {
  const wizard = getWizardState(setupModel);
  const grouped = new Map<string, SetupFieldConfig[]>();
  const fallbackGroupId =
    wizard.groups.find((group) => group.sectionId === 'sonstiges')?.sectionId ||
    wizard.groups[wizard.groups.length - 1]?.sectionId ||
    'sonstiges';

  for (const group of wizard.groups) {
    grouped.set(group.sectionId, []);
  }

  for (const field of fields) {
    let sectionId = wizard.assignments[field.fieldId];
    if (!sectionId && wizard.deferredFieldIds.includes(field.fieldId)) {
      sectionId = fallbackGroupId;
    }
    if (!sectionId) continue;

    const bucket = grouped.get(sectionId) || [];
    bucket.push(
      fieldConfigFromDetected(
        field,
        findExistingFieldConfig(setupModel, field.fieldId) || undefined
      )
    );
    grouped.set(sectionId, bucket);
  }

  const singleSections = wizard.groups
    .map((group) => ({
      sectionId: group.sectionId,
      label: group.label,
      fields: grouped.get(group.sectionId) || []
    }))
    .filter((section) => section.fields.length > 0);

  if (singleSections.length === 0) {
    singleSections.push({
      sectionId: fallbackGroupId,
      label: 'Sonstiges',
      fields: []
    });
  }

  const tableSections = buildTableSectionsFromWizard(wizard, fields);
  const nextModel = {
    ...setupModel,
    single_sections: singleSections,
    table_sections: tableSections,
    wizard: {
      ...wizard,
      step: 'fields' as const
    },
    updatedAt: nowIso()
  };
  return {
    ...nextModel,
    section_order: syncSectionOrder({
      ...nextModel,
      section_order: buildLegacySectionOrder(nextModel)
    })
  };
}

export function listSetupSections(setupModel: Record<string, unknown>): Array<{
  sectionId: string;
  label: string;
  fields: SetupFieldConfig[];
}> {
  const singleSections = Array.isArray(setupModel.single_sections) ? setupModel.single_sections : [];
  const byId = new Map(
    singleSections.map((section) => [
      String(section?.sectionId || ''),
      {
        sectionId: String(section?.sectionId || ''),
        label: String(section?.label || 'Gruppe'),
        fields: (Array.isArray(section?.fields) ? section.fields : []) as SetupFieldConfig[]
      }
    ])
  );

  const orderedIds = (Array.isArray(setupModel.section_order) ? setupModel.section_order : [])
    .filter((entry) => String(entry?.kind || '') === 'single')
    .map((entry) => String(entry?.id || ''))
    .filter(Boolean);

  const seen = new Set<string>();
  const sections = [];
  for (const sectionId of orderedIds) {
    const section = byId.get(sectionId);
    if (!section || seen.has(sectionId)) continue;
    seen.add(sectionId);
    sections.push(section);
  }
  for (const [sectionId, section] of byId) {
    if (!seen.has(sectionId)) sections.push(section);
  }
  return sections;
}

export function updateSetupField(
  setupModel: Record<string, unknown>,
  sectionId: string,
  fieldId: string,
  patch: Partial<SetupFieldConfig>
): Record<string, unknown> {
  const singleSections = Array.isArray(setupModel.single_sections)
    ? [...setupModel.single_sections]
    : [];

  const nextSections = singleSections.map((section) => {
    if (String(section?.sectionId) !== String(sectionId)) return section;
    const fields = Array.isArray(section?.fields) ? [...section.fields] : [];
    return {
      ...section,
      fields: fields.map((field) =>
        String(field?.fieldId) === String(fieldId) ? { ...field, ...patch } : field
      )
    };
  });

  return {
    ...setupModel,
    single_sections: nextSections,
    updatedAt: nowIso()
  };
}

/** PDF page height assumption for A4-like templates (points, origin bottom-left). */
const PDF_PAGE_HEIGHT = 842;
const PDF_PAGE_WIDTH = 595;

/**
 * Place group-selection overlay away from the active field so it stays visible.
 * PDF Y increases upward — high centerY = field near top of page on screen.
 */
export function resolveOverlayPlacement(rect: number[] | null): OverlayPlacement {
  if (!rect || rect.length < 4) return 'bottom';
  const [x1, y1, x2, y2] = rect;
  const centerX = (x1 + x2) / 2;
  const centerY = (y1 + y2) / 2;

  const topThird = PDF_PAGE_HEIGHT * 0.68;
  const bottomThird = PDF_PAGE_HEIGHT * 0.32;
  const leftThird = PDF_PAGE_WIDTH * 0.33;
  const rightThird = PDF_PAGE_WIDTH * 0.67;

  // Field near top of page → selection panel at bottom of preview.
  if (centerY >= topThird) return 'bottom';
  // Field near bottom of page → selection panel at top.
  if (centerY <= bottomThird) return 'top';
  // Field on left edge → panel on right.
  if (centerX <= leftThird) return 'right';
  // Field on right edge → panel on left.
  if (centerX >= rightThird) return 'left';
  return 'bottom';
}

export function resolveCurrentMappingIndex(
  fields: MappingField[],
  wizard: SetupWizardState
): number {
  if (fields.length === 0) return 0;
  const stored = Math.max(0, Number(wizard.currentFieldIndex || 0));
  return Math.min(stored, fields.length - 1);
}

/** Walkthrough helper: jump to the next open field when the stored one is already handled. */
export function resolveWalkthroughMappingIndex(
  fields: MappingField[],
  wizard: SetupWizardState
): number {
  if (fields.length === 0) return 0;
  const stored = resolveCurrentMappingIndex(fields, wizard);
  const storedField = fields[stored];
  if (
    storedField &&
    !wizard.assignments[storedField.fieldId] &&
    !wizard.tableAssignments[storedField.fieldId] &&
    !wizard.deferredFieldIds.includes(storedField.fieldId)
  ) {
    return stored;
  }
  const next = getNextUnassignedIndex(fields, wizard, stored + 1);
  return next >= 0 ? next : stored;
}

export function templateDisplayStatus(
  templateStatus: string,
  isActive: boolean
): { label: string; tone: 'neutral' | 'success' | 'warning' | 'danger' | 'info' | 'accent' } {
  if (isActive && templateStatus === 'ready') {
    return { label: 'Aktiv', tone: 'success' };
  }
  if (templateStatus === 'ready') {
    return { label: 'Ready', tone: 'info' };
  }
  if (templateStatus === 'in_progress') {
    return { label: 'In Bearbeitung', tone: 'warning' };
  }
  if (templateStatus === 'archived') {
    return { label: 'Archiviert', tone: 'neutral' };
  }
  return { label: 'Draft', tone: 'accent' };
}

export function ensureWizardInitialized(setupModel: Record<string, unknown>): Record<string, unknown> {
  if (setupModel.wizard && typeof setupModel.wizard === 'object') {
    return setupModel;
  }
  return withWizardState(setupModel, { step: 'structure', groups: [], tables: [], structure: [] });
}

export function hasTableSections(setupModel: Record<string, unknown>): boolean {
  return Array.isArray(setupModel.table_sections) && setupModel.table_sections.length > 0;
}

export function shouldSkipSetupIntros(setupModel: Record<string, unknown>): boolean {
  const wizard = getWizardState(setupModel);
  return wizard.setupCompleted === true || wizard.editMode === true;
}

export function markSetupCompleted(setupModel: Record<string, unknown>): Record<string, unknown> {
  return withWizardState(setupModel, {
    setupCompleted: true,
    editMode: false,
    step: 'fields'
  });
}

export function enterEditMode(
  setupModel: Record<string, unknown>,
  step: SetupWizardStep
): Record<string, unknown> {
  return withWizardState(setupModel, { editMode: true, step });
}

export function exitEditMode(setupModel: Record<string, unknown>): Record<string, unknown> {
  return withWizardState(setupModel, { editMode: false, step: 'fields' });
}

export function resolveTemplateDetailPath(templateId: string): Href {
  return `/bautagebuch/setup/${templateId}/detail` as Href;
}

export function resolveTemplateEditPath(templateId: string): Href {
  return `/bautagebuch/setup/${templateId}/edit` as Href;
}

export function resolveSetupEditStepPath(templateId: string, step: SetupWizardStep): Href {
  if (step === 'structure') {
    return `/bautagebuch/setup/${templateId}/mapping` as Href;
  }
  if (step === 'assign') {
    return `/bautagebuch/setup/${templateId}/assign` as Href;
  }
  return `/bautagebuch/setup/${templateId}/fields` as Href;
}

export function resolveTemplateOpenPath(
  templateId: string,
  setupModel: Record<string, unknown>,
  templateStatus: string,
  templateKind = ''
): Href {
  if (templateKind === 'builtin-etb') {
    return `/bautagebuch/setup/${templateId}/fields` as Href;
  }
  if (templateStatus === 'ready' || templateStatus === 'archived') {
    return resolveTemplateDetailPath(templateId);
  }
  return resolveSetupEntryPath(templateId, setupModel, templateKind);
}

export function countAssignedFieldsForStructureItem(
  setupModel: Record<string, unknown>,
  item: { id: string; type: 'group' | 'table' }
): number {
  const wizard = getWizardState(setupModel);
  if (item.type === 'group') {
    const fromWizard = Object.values(wizard.assignments).filter(
      (sectionId) => sectionId === item.id
    ).length;
    const section = listSetupSections(setupModel).find((entry) => entry.sectionId === item.id);
    const fromSection = (section?.fields || []).filter((field) => field.skipped !== true).length;
    return Math.max(fromWizard, fromSection);
  }
  const fromWizard = Object.values(wizard.tableAssignments).filter(
    (assignment) => assignment.tableId === item.id
  ).length;
  const table = (Array.isArray(setupModel.table_sections) ? setupModel.table_sections : []).find(
    (entry) => String(entry?.tableId || '') === item.id
  );
  const fromCells = (Array.isArray(table?.rows) ? table.rows : []).reduce(
    (sum: number, row: { cells?: Array<{ fieldId?: string }> }) => {
      const cells = Array.isArray(row?.cells) ? row.cells : [];
      return sum + cells.filter((cell: { fieldId?: string }) => String(cell?.fieldId || '').trim()).length;
    },
    0
  );
  return Math.max(fromWizard, fromCells);
}

export function shouldShowStructureIntro(setupModel: Record<string, unknown>): boolean {
  if (shouldSkipSetupIntros(setupModel)) return false;
  const wizard = getWizardState(setupModel);
  if (wizard.step !== 'structure') return false;
  if (wizard.structureIntroSeen) return false;
  if (wizard.structure.length > 0) return false;
  return true;
}

export function markStructureIntroSeen(setupModel: Record<string, unknown>): Record<string, unknown> {
  return withWizardState(setupModel, { structureIntroSeen: true });
}

export function shouldShowAssignIntro(setupModel: Record<string, unknown>): boolean {
  if (shouldSkipSetupIntros(setupModel)) return false;
  const wizard = getWizardState(setupModel);
  if (wizard.step !== 'assign') return false;
  if (wizard.assignIntroSeen) return false;
  const hasAssignments = Object.keys(wizard.assignments).length > 0;
  const hasTableAssignments = Object.keys(wizard.tableAssignments).length > 0;
  const hasDeferred = wizard.deferredFieldIds.length > 0;
  if (hasAssignments || hasTableAssignments || hasDeferred) return false;
  return true;
}

export function markAssignIntroSeen(setupModel: Record<string, unknown>): Record<string, unknown> {
  return withWizardState(setupModel, { assignIntroSeen: true });
}

export function shouldShowFieldsIntro(
  setupModel: Record<string, unknown>,
  templateKind = ''
): boolean {
  if (templateKind === 'builtin-etb') return false;
  if (shouldSkipSetupIntros(setupModel)) return false;
  const wizard = getWizardState(setupModel);
  if (wizard.step !== 'fields') return false;
  if (wizard.fieldsIntroSeen) return false;
  return true;
}

export function markFieldsIntroSeen(setupModel: Record<string, unknown>): Record<string, unknown> {
  return withWizardState(setupModel, { fieldsIntroSeen: true });
}

export function resolveFieldsSetupPath(
  templateId: string,
  setupModel: Record<string, unknown>,
  templateKind = ''
): Href {
  if (shouldShowFieldsIntro(setupModel, templateKind)) {
    return `/bautagebuch/setup/${templateId}/fields-intro` as Href;
  }
  return `/bautagebuch/setup/${templateId}/fields` as Href;
}

export function resolveSetupEntryPath(
  templateId: string,
  setupModel: Record<string, unknown>,
  templateKind = ''
): Href {
  if (templateKind === 'builtin-etb') {
    return `/bautagebuch/setup/${templateId}/fields` as Href;
  }
  const wizard = getWizardState(setupModel);
  if (wizard.step === 'fields') {
    return resolveFieldsSetupPath(templateId, setupModel, templateKind);
  }
  if (hasTableSections(setupModel)) {
    return `/bautagebuch/setup/${templateId}/fields` as Href;
  }
  if (wizard.step === 'assign') {
    if (shouldShowAssignIntro(setupModel)) {
      return `/bautagebuch/setup/${templateId}/assign-intro` as Href;
    }
    return `/bautagebuch/setup/${templateId}/assign` as Href;
  }
  if (shouldShowStructureIntro(setupModel)) {
    return `/bautagebuch/setup/${templateId}/intro` as Href;
  }
  return `/bautagebuch/setup/${templateId}/mapping` as Href;
}

export type MappingCompletionSummary = {
  totalFields: number;
  groupCount: number;
  groups: Array<{ sectionId: string; label: string; fieldCount: number }>;
  unassignedCount: number;
  assignedCount: number;
};

export function getMappingCompletionSummary(
  setupModel: Record<string, unknown>,
  mappingFields: MappingField[]
): MappingCompletionSummary {
  const wizard = getWizardState(setupModel);
  const groups = wizard.groups.map((group) => ({
    sectionId: group.sectionId,
    label: group.label,
    fieldCount: mappingFields.filter(
      (field) => wizard.assignments[field.fieldId] === group.sectionId
    ).length
  }));

  return {
    totalFields: mappingFields.length,
    groupCount: wizard.groups.length,
    groups,
    unassignedCount: wizard.deferredFieldIds.length,
    assignedCount: Object.keys(wizard.assignments).length
  };
}

export type MappingTransitionCheck = {
  hasGroups: boolean;
  hasAssignedFields: boolean;
  unassignedCount: number;
  issues: string[];
};

export function checkMappingTransition(
  setupModel: Record<string, unknown>,
  mappingFields: MappingField[]
): MappingTransitionCheck {
  const wizard = getWizardState(setupModel);
  const hasGroups = wizard.groups.length > 0;
  const hasAssignedFields = Object.keys(wizard.assignments).length > 0;
  const unassignedCount = wizard.deferredFieldIds.length;
  const issues: string[] = [];

  if (!hasGroups) {
    issues.push('Es wurden noch keine Gruppen erstellt.');
  }
  if (mappingFields.length > 0 && !hasAssignedFields && unassignedCount === 0) {
    issues.push('Es wurden noch keine Felder Gruppen zugeordnet.');
  }
  if (unassignedCount > 0) {
    issues.push(`${unassignedCount} Felder wurden keiner Gruppe zugeordnet.`);
  }

  return { hasGroups, hasAssignedFields, unassignedCount, issues };
}
