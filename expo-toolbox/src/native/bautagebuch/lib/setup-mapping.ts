import type { Href } from 'expo-router';

import type { DetectedField, SetupFieldConfig, SetupWizardGroup, SetupWizardState } from '../types';

export type MappingField = {
  fieldId: string;
  fieldName: string;
  labelCandidate: string;
  type: string;
  page: number;
  orderIndex: number;
  rect: number[] | null;
};

export type MappingProgress = {
  current: number;
  total: number;
  percent: number;
  remaining: number;
  assigned: number;
};

export type OverlayPlacement = 'top' | 'bottom' | 'left' | 'right';

function nowIso(): string {
  return new Date().toISOString();
}

function createId(prefix = 'id'): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
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
    .map((field) => ({
      fieldId: String(field.fieldId),
      fieldName: String(field.fieldName || ''),
      labelCandidate: String(field.labelCandidate || field.fieldName || 'Feld'),
      type: String(field.type || 'text'),
      page: Number(field.page || 1),
      orderIndex: Number(field.orderIndex || 0),
      rect: Array.isArray(field.rect) ? field.rect.slice(0, 4) : null
    }));
}

export function getWizardState(setupModel: Record<string, unknown>): SetupWizardState {
  const raw = (setupModel.wizard || {}) as Partial<SetupWizardState>;
  const groups =
    Array.isArray(raw.groups) && raw.groups.length > 0
      ? raw.groups.map((group) => ({
          sectionId: String(group.sectionId || ''),
          label: String(group.label || 'Gruppe')
        }))
      : [];

  return {
    step: raw.step === 'fields' ? 'fields' : 'mapping',
    currentFieldIndex: Math.max(0, Number(raw.currentFieldIndex || 0)),
    groups,
    assignments:
      raw.assignments && typeof raw.assignments === 'object'
        ? Object.fromEntries(
            Object.entries(raw.assignments).map(([fieldId, sectionId]) => [
              String(fieldId),
              String(sectionId)
            ])
          )
        : {},
    deferredFieldIds: Array.isArray(raw.deferredFieldIds)
      ? raw.deferredFieldIds.map((fieldId) => String(fieldId))
      : []
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
      assignments: wizard.assignments || current.assignments,
      deferredFieldIds: wizard.deferredFieldIds || current.deferredFieldIds
    },
    updatedAt: nowIso()
  };
}

export function getMappingProgress(
  fields: MappingField[],
  wizard: SetupWizardState
): MappingProgress {
  const total = fields.length;
  const assigned = fields.filter((field) => Boolean(wizard.assignments[field.fieldId])).length;
  const handled = fields.filter(
    (field) =>
      Boolean(wizard.assignments[field.fieldId]) || wizard.deferredFieldIds.includes(field.fieldId)
  ).length;
  const current = total > 0 ? Math.min(total, Math.max(1, handled + 1)) : 0;
  const percent = total > 0 ? Math.round((handled / total) * 100) : 100;
  return {
    current,
    total,
    percent,
    remaining: Math.max(0, total - handled),
    assigned
  };
}

export function getNextUnassignedIndex(
  fields: MappingField[],
  wizard: SetupWizardState,
  startIndex = 0
): number {
  for (let index = Math.max(0, startIndex); index < fields.length; index += 1) {
    const field = fields[index];
    if (!wizard.assignments[field.fieldId] && !wizard.deferredFieldIds.includes(field.fieldId)) {
      return index;
    }
  }
  return -1;
}

export function isMappingComplete(fields: MappingField[], wizard: SetupWizardState): boolean {
  if (fields.length === 0) return true;
  return fields.every(
    (field) =>
      Boolean(wizard.assignments[field.fieldId]) || wizard.deferredFieldIds.includes(field.fieldId)
  );
}

export function assignFieldToGroup(
  setupModel: Record<string, unknown>,
  fieldId: string,
  sectionId: string
): Record<string, unknown> {
  const wizard = getWizardState(setupModel);
  const assignments = { ...wizard.assignments, [fieldId]: sectionId };
  const deferredFieldIds = wizard.deferredFieldIds.filter((entry) => entry !== fieldId);
  return withWizardState(setupModel, { assignments, deferredFieldIds });
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
  return withWizardState(setupModel, { assignments, deferredFieldIds });
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
  existing?: SetupFieldConfig
): SetupFieldConfig {
  return {
    fieldId: field.fieldId,
    fieldName: field.fieldName,
    label: String(existing?.label || field.labelCandidate || field.fieldName || 'Feld').trim(),
    type: field.type,
    required: Boolean(existing?.required),
    skipped: Boolean(existing?.skipped),
    multiline: Boolean(existing?.multiline),
    defaultValue: String(existing?.defaultValue || ''),
    hint: String(existing?.hint || ''),
    page: field.page,
    rect: field.rect
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

  return {
    ...setupModel,
    single_sections: singleSections,
    wizard: {
      ...wizard,
      step: 'fields' as const
    },
    updatedAt: nowIso()
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
  const storedField = fields[stored];
  if (
    storedField &&
    !wizard.assignments[storedField.fieldId] &&
    !wizard.deferredFieldIds.includes(storedField.fieldId)
  ) {
    return stored;
  }
  const next = getNextUnassignedIndex(fields, wizard, 0);
  return next >= 0 ? next : Math.min(stored, fields.length - 1);
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
  return withWizardState(setupModel, { step: 'mapping', groups: [] });
}

export function hasTableSections(setupModel: Record<string, unknown>): boolean {
  return Array.isArray(setupModel.table_sections) && setupModel.table_sections.length > 0;
}

export function resolveSetupEntryPath(
  templateId: string,
  setupModel: Record<string, unknown>,
  templateKind = ''
): Href {
  if (templateKind === 'builtin-etb' || hasTableSections(setupModel)) {
    return `/bautagebuch/setup/${templateId}/fields` as Href;
  }
  const wizard = getWizardState(setupModel);
  if (wizard.step === 'fields') {
    return `/bautagebuch/setup/${templateId}/fields` as Href;
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
