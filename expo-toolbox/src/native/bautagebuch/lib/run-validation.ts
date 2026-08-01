import { buildRunSections, inputKeyForField, requiredMissingCount } from './setup-model.js';

const GEWERK_FIELD_NAMES = new Set(['Text3', 'Text5', 'Text6', 'Text7', 'Text8']);
const SHIFT_FIELD_NAMES = new Set(['Check Box1', 'Check Box2', 'Check Box3']);
const PHOTO_DOC_SECTION_ID = 'photo-doc';

export type RunSection =
  | ReturnType<typeof buildRunSections>[number]
  | {
      sectionId: string;
      kind: 'photo-doc';
      label: string;
    };

function tableRowCountKey(tableId: string): string {
  return `__tableRows:${String(tableId || '').trim()}`;
}

function cellHasValue(cell: { type?: string }, values: Record<string, unknown>): boolean {
  const cellId = String((cell as { cellId?: string })?.cellId || '').trim();
  if (!cellId) return false;
  const key = `cell:${cellId}`;
  const value = values[key];
  if (value === true) return true;
  return String(value ?? '').trim().length > 0;
}

function detectedVisibleRowCountFromValues(section: RunSection, values: Record<string, unknown>): number {
  if (!section || section.kind !== 'table') return 0;
  let lastFilled = 0;
  const rows = section.rows || [];
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    const rowHasValue = (row.cells || []).some((cell: { cellId?: string; type?: string }) => cellHasValue(cell, values));
    if (rowHasValue) lastFilled = index + 1;
  }
  return lastFilled;
}

export function configuredTableRowCount(
  setupModel: Record<string, unknown> | null | undefined,
  tableId: string
): number {
  const tables = setupModel?.table_sections;
  if (!Array.isArray(tables)) return 0;
  const table = tables.find((entry) => String((entry as { tableId?: string })?.tableId || '') === String(tableId));
  const rows = Array.isArray((table as { rows?: unknown[] })?.rows)
    ? (table as { rows: unknown[] }).rows
    : [];
  return rows.length;
}

export function visibleRowCountForSection(
  section: RunSection,
  values: Record<string, unknown>,
  options: { maxRows?: number } = {}
): number {
  if (!section || section.kind !== 'table') return 0;
  const rows = Array.isArray(section.rows) ? section.rows : [];
  if (rows.length === 0) return 0;

  const configuredMaxRows = Number(options.maxRows);
  const maxRows = Math.max(
    rows.length,
    Number.isFinite(configuredMaxRows) ? Math.floor(configuredMaxRows) : 0
  );
  const storedCount = Number(values[tableRowCountKey(String(section.tableId || ''))]);
  const detectedCount = detectedVisibleRowCountFromValues(section, values);
  const minimum = Math.max(1, detectedCount);

  if (!Number.isFinite(storedCount) || storedCount <= 0) return Math.min(maxRows, minimum);
  return Math.min(maxRows, Math.max(minimum, Math.floor(storedCount)));
}

function isGewerkGroupSection(section: RunSection): boolean {
  if (!section || section.kind !== 'single') return false;
  const sectionId = String(section.sectionId || '').toLowerCase();
  const label = String(section.label || '').toLowerCase();
  return sectionId.includes('header') || label.includes('kopfdaten');
}

export function requiredAnyGroupsForSection(section: RunSection) {
  if (!section || section.kind !== 'single' || !isGewerkGroupSection(section)) return [];

  const groups: Array<{ groupId: string; label: string; fieldIds: string[] }> = [];
  const gewerkFieldIds = (section.fields || [])
    .filter((field: { fieldName?: string }) => GEWERK_FIELD_NAMES.has(String(field.fieldName || '').trim()))
    .map((field: { fieldId?: string }) => String(field.fieldId || '').trim())
    .filter(Boolean);
  if (gewerkFieldIds.length > 0) {
    groups.push({
      groupId: `gewerk:${section.sectionId || 'single'}`,
      label: 'Gewerk',
      fieldIds: gewerkFieldIds
    });
  }

  const shiftFieldIds = (section.fields || [])
    .filter((field: { fieldName?: string }) => SHIFT_FIELD_NAMES.has(String(field.fieldName || '').trim()))
    .map((field: { fieldId?: string }) => String(field.fieldId || '').trim())
    .filter(Boolean);
  if (shiftFieldIds.length > 0) {
    groups.push({
      groupId: `shift:${section.sectionId || 'single'}`,
      label: 'Schicht',
      fieldIds: shiftFieldIds
    });
  }

  return groups;
}

export function sectionRunOptions(
  section: RunSection,
  values: Record<string, unknown> = {},
  options: { maxRows?: number } = {}
) {
  if (!section) return {};
  if (section.kind === 'table') {
    return { visibleRowCount: visibleRowCountForSection(section, values, options) };
  }
  if (section.kind === 'single') {
    return { requiredAnyGroups: requiredAnyGroupsForSection(section) };
  }
  return {};
}

export function runSectionMissingCount(
  section: RunSection,
  values: Record<string, unknown>,
  options: { maxRows?: number } = {}
): number {
  if (!section) return 0;
  if (section.kind === 'photo-doc' || section.sectionId === PHOTO_DOC_SECTION_ID) {
    return 0;
  }
  const runOptions =
    section.kind === 'table' ? sectionRunOptions(section, values, options) : sectionRunOptions(section, values);
  return requiredMissingCount(section, values, runOptions);
}

export function isPhotoDocRequiredMissing(enabled: boolean | null | undefined): boolean {
  return enabled !== true && enabled !== false;
}

export function buildRunSectionsWithPhotoDoc(setupModel: Record<string, unknown>): RunSection[] {
  const base = buildRunSections(setupModel) as RunSection[];
  return [
    ...base,
    {
      sectionId: PHOTO_DOC_SECTION_ID,
      kind: 'photo-doc',
      label: 'Fotodokumentation'
    }
  ];
}

export function computeTotalMissingRequired(
  setupModel: Record<string, unknown>,
  values: Record<string, unknown>,
  photoDocEnabled: boolean | null | undefined
): number {
  const sections = buildRunSectionsWithPhotoDoc(setupModel);
  return sections.reduce((sum, section) => {
    if (section.kind === 'photo-doc' || section.sectionId === PHOTO_DOC_SECTION_ID) {
      return sum + (isPhotoDocRequiredMissing(photoDocEnabled) ? 1 : 0);
    }
    const rowLimit =
      section.kind === 'table'
        ? { maxRows: configuredTableRowCount(setupModel, String(section.tableId || '')) }
        : {};
    return sum + runSectionMissingCount(section, values, rowLimit);
  }, 0);
}

export function exportBlockedMessage(missingCount: number): string {
  if (missingCount <= 0) return '';
  return `Export blockiert: ${missingCount} Pflichtfeld${missingCount === 1 ? '' : 'er'} fehlen.`;
}

export function fieldKeyForId(fieldId: string): string {
  return inputKeyForField({ fieldId } as { fieldId: string });
}
