import type {
  SetupStructureGroup,
  SetupStructureItem,
  SetupStructureTable,
  SetupStructureTableColumn,
  SetupWizardGroup,
  SetupWizardTable
} from '../types';
import { getWizardState, withWizardState } from './setup-mapping';

function createId(prefix = 'id'): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

function nowIso(): string {
  return new Date().toISOString();
}

function normalizeOrder(items: SetupStructureItem[]): SetupStructureItem[] {
  return items
    .slice()
    .sort((left, right) => left.order - right.order)
    .map((item, index) => ({ ...item, order: index }));
}

function renumberStructure(items: SetupStructureItem[]): SetupStructureItem[] {
  return items.map((item, index) => ({ ...item, order: index }));
}

export function migrateStructureFromLegacy(
  groups: SetupWizardGroup[],
  tables: SetupWizardTable[]
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

export function getStructureItems(setupModel: Record<string, unknown>): SetupStructureItem[] {
  const wizard = getWizardState(setupModel);
  const raw = (wizard as { structure?: SetupStructureItem[] }).structure;
  if (Array.isArray(raw) && raw.length > 0) {
    return normalizeOrder(
      raw.map((item) => {
        if (item.type === 'table') {
          return {
            ...item,
            columns: [...(item.columns || [])].sort((a, b) => a.order - b.order)
          };
        }
        return item;
      })
    );
  }
  if (wizard.groups.length > 0 || wizard.tables.length > 0) {
    return migrateStructureFromLegacy(wizard.groups, wizard.tables);
  }
  return [];
}

function structureToGroupsTables(structure: SetupStructureItem[]): {
  groups: SetupWizardGroup[];
  tables: SetupWizardTable[];
} {
  const groups: SetupWizardGroup[] = [];
  const tables: SetupWizardTable[] = [];
  for (const item of normalizeOrder(structure)) {
    if (item.type === 'group') {
      groups.push({
        sectionId: item.id,
        label: item.name,
        description: item.description
      });
      continue;
    }
    tables.push({
      tableId: item.id,
      label: item.name,
      rowCount: 1,
      columns: item.columns.map((column) => ({
        columnId: column.id,
        label: column.name,
        type: 'text' as const,
        required: false,
        multiline: false,
        skipped: false
      }))
    });
  }
  return { groups, tables };
}

function buildSectionShellsFromStructure(structure: SetupStructureItem[]) {
  const single_sections = normalizeOrder(structure)
    .filter((item): item is SetupStructureGroup => item.type === 'group')
    .map((group) => ({
      sectionId: group.id,
      label: group.name,
      description: group.description || '',
      fields: [] as unknown[]
    }));

  const table_sections = normalizeOrder(structure)
    .filter((item): item is SetupStructureTable => item.type === 'table')
    .map((table) => {
      const columns = table.columns.map((column) => ({
        columnId: column.id,
        label: column.name,
        type: 'text' as const,
        required: false,
        multiline: false,
        skipped: false
      }));
      return {
        tableId: table.id,
        label: table.name,
        columns,
        rows: [
          {
            rowId: 'row_1',
            index: 1,
            cells: columns.map((column) => ({
              cellId: `${table.id}_row_1_${column.columnId}`,
              tableId: table.id,
              rowId: 'row_1',
              columnId: column.columnId,
              fieldId: '',
              fieldName: '',
              label: column.label,
              type: 'text',
              options: [],
              page: 1,
              rect: null,
              skipped: true,
              required: false,
              multiline: false
            }))
          }
        ]
      };
    });

  const section_order = normalizeOrder(structure).map((item) =>
    item.type === 'group' ? { kind: 'single' as const, id: item.id } : { kind: 'table' as const, id: item.id }
  );

  return { single_sections, table_sections, section_order };
}

export function withStructureItems(
  setupModel: Record<string, unknown>,
  structure: SetupStructureItem[]
): Record<string, unknown> {
  const normalized = renumberStructure(structure);
  const { groups, tables } = structureToGroupsTables(normalized);
  return withWizardState(setupModel, {
    structure: normalized,
    groups,
    tables
  } as Partial<ReturnType<typeof getWizardState>>);
}

export function addStructureGroup(
  setupModel: Record<string, unknown>,
  input: { name: string; description?: string }
): Record<string, unknown> {
  const structure = getStructureItems(setupModel);
  const item: SetupStructureGroup = {
    id: createId('group'),
    name: String(input.name || '').trim() || 'Neue Gruppe',
    description: String(input.description || '').trim() || undefined,
    type: 'group',
    order: structure.length
  };
  return withStructureItems(setupModel, [...structure, item]);
}

export function addStructureTable(
  setupModel: Record<string, unknown>,
  input: { name: string; columns: Array<{ name: string }> }
): Record<string, unknown> {
  const structure = getStructureItems(setupModel);
  const columns: SetupStructureTableColumn[] = (input.columns || [])
    .map((column, index) => ({
      id: createId('col'),
      name: String(column.name || '').trim() || `Spalte ${index + 1}`,
      order: index
    }))
    .filter((column) => column.name.length > 0);
  const item: SetupStructureTable = {
    id: createId('table'),
    name: String(input.name || '').trim() || 'Neue Tabelle',
    type: 'table',
    order: structure.length,
    columns
  };
  return withStructureItems(setupModel, [...structure, item]);
}

export function updateStructureGroup(
  setupModel: Record<string, unknown>,
  id: string,
  patch: { name?: string; description?: string }
): Record<string, unknown> {
  const structure = getStructureItems(setupModel).map((item) => {
    if (item.type !== 'group' || item.id !== id) return item;
    return {
      ...item,
      name: patch.name !== undefined ? String(patch.name).trim() || item.name : item.name,
      description:
        patch.description !== undefined ? String(patch.description).trim() || undefined : item.description
    };
  });
  return withStructureItems(setupModel, structure);
}

export function updateStructureTable(
  setupModel: Record<string, unknown>,
  id: string,
  patch: { name?: string; columns?: Array<{ id?: string; name: string }> }
): Record<string, unknown> {
  const structure = getStructureItems(setupModel).map((item) => {
    if (item.type !== 'table' || item.id !== id) return item;
    if (!patch.columns) {
      return {
        ...item,
        name: patch.name !== undefined ? String(patch.name).trim() || item.name : item.name
      };
    }
    return {
      ...item,
      name: patch.name !== undefined ? String(patch.name).trim() || item.name : item.name,
      columns: patch.columns.map((column, index) => ({
        id: column.id || createId('col'),
        name: String(column.name || '').trim() || `Spalte ${index + 1}`,
        order: index
      }))
    };
  });
  return withStructureItems(setupModel, structure);
}

export function deleteStructureItem(setupModel: Record<string, unknown>, id: string): Record<string, unknown> {
  const structure = getStructureItems(setupModel).filter((item) => item.id !== id);
  return withStructureItems(setupModel, structure);
}

export function moveStructureItem(
  setupModel: Record<string, unknown>,
  id: string,
  direction: -1 | 1
): Record<string, unknown> {
  const structure = normalizeOrder(getStructureItems(setupModel));
  const index = structure.findIndex((item) => item.id === id);
  if (index < 0) return setupModel;
  const target = index + direction;
  if (target < 0 || target >= structure.length) return setupModel;
  const next = [...structure];
  [next[index], next[target]] = [next[target], next[index]];
  return withStructureItems(setupModel, next);
}

export function validateStructureStep(structure: SetupStructureItem[]): string | null {
  if (structure.length === 0) {
    return 'Lege mindestens eine Gruppe oder Tabelle an.';
  }
  for (const item of structure) {
    if (item.type === 'table' && item.columns.length === 0) {
      return `Tabelle „${item.name}“ benötigt mindestens eine Spalte.`;
    }
  }
  return null;
}

export function completeStructureStep(setupModel: Record<string, unknown>): Record<string, unknown> {
  const structure = getStructureItems(setupModel);
  const issue = validateStructureStep(structure);
  if (issue) {
    throw new Error(issue);
  }
  const { groups, tables } = structureToGroupsTables(structure);
  const shells = buildSectionShellsFromStructure(structure);
  const wizard = getWizardState(setupModel);
  return {
    ...setupModel,
    ...shells,
    updatedAt: nowIso(),
    wizard: {
      ...wizard,
      step: 'assign' as const,
      structure,
      groups,
      tables,
      assignments: {},
      tableAssignments: {},
      deferredFieldIds: [],
      fieldLabels: {},
      currentFieldIndex: 0
    }
  };
}
