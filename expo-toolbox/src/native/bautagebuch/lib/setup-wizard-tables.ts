import type { SetupWizardTable, SetupWizardTableColumn, SetupWizardState } from '../types';
import type { MappingField } from './setup-mapping';

function tableCellEntry(
  table: SetupWizardTable,
  rowIndex: number,
  column: SetupWizardTableColumn,
  field?: MappingField
) {
  const rowId = `row_${rowIndex + 1}`;
  const cellId = `${table.tableId}_${rowId}_${column.columnId}`;
  if (!field) {
    return {
      cellId,
      tableId: table.tableId,
      rowId,
      columnId: column.columnId,
      fieldId: '',
      fieldName: '',
      label: column.label,
      type: column.type,
      options: [],
      page: 1,
      rect: null,
      skipped: true,
      required: column.required === true,
      multiline: column.multiline === true
    };
  }

  return {
    cellId,
    tableId: table.tableId,
    rowId,
    columnId: column.columnId,
    fieldId: field.fieldId,
    fieldName: field.fieldName,
    label: String(field.labelCandidate || field.fieldName || column.label).trim(),
    type: field.type || column.type,
    options: Array.isArray(field.options) ? field.options : [],
    page: field.page,
    rect: field.rect,
    skipped: false,
    required: column.required === true,
    multiline: column.multiline === true
  };
}

export function buildTableSectionsFromWizard(
  wizard: SetupWizardState,
  fields: MappingField[]
): Record<string, unknown>[] {
  const fieldById = new Map(fields.map((field) => [field.fieldId, field]));

  return wizard.tables
    .filter((table) => table.columns.length > 0 && table.rowCount > 0)
    .map((table) => {
      const columns = table.columns.map((column) => ({
        columnId: column.columnId,
        label: column.label,
        type: column.type,
        required: column.required === true,
        skipped: column.skipped === true,
        multiline: column.multiline === true
      }));

      const rows = Array.from({ length: table.rowCount }, (_, rowIndex) => {
        const rowId = `row_${rowIndex + 1}`;
        const cells = table.columns.map((column) => {
          const assignedFieldId = Object.entries(wizard.tableAssignments).find(
            ([, assignment]) =>
              assignment.tableId === table.tableId &&
              assignment.rowIndex === rowIndex &&
              assignment.columnId === column.columnId
          )?.[0];
          const field = assignedFieldId ? fieldById.get(assignedFieldId) : undefined;
          const customLabel = assignedFieldId ? wizard.fieldLabels?.[assignedFieldId] : undefined;
          const cell = tableCellEntry(table, rowIndex, column, field);
          if (customLabel && assignedFieldId) {
            return { ...cell, label: customLabel };
          }
          return cell;
        });
        return {
          rowId,
          index: rowIndex + 1,
          cells
        };
      });

      return {
        tableId: table.tableId,
        label: table.label,
        columns,
        rows
      };
    });
}

export function tableAssignmentKey(assignment: {
  tableId: string;
  rowIndex: number;
  columnId: string;
}): string {
  return `${assignment.tableId}:${assignment.rowIndex}:${assignment.columnId}`;
}

export function findFieldAssignedToTableCell(
  wizard: SetupWizardState,
  assignment: { tableId: string; rowIndex: number; columnId: string }
): string | null {
  for (const [fieldId, entry] of Object.entries(wizard.tableAssignments)) {
    if (tableAssignmentKey(entry) === tableAssignmentKey(assignment)) {
      return fieldId;
    }
  }
  return null;
}
