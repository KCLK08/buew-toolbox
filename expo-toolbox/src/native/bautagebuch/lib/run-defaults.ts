import { inputKeyForField } from './setup-model.js';

const RUN_DEFAULTS_BY_FIELD_NAME = new Map<string, string>([['Text2', 'Kazim Celik']]);
const DATE_FIELD_NAME_PATTERN = /^Date\d+$/i;

function todayDateLabel(): string {
  const now = new Date();
  const day = String(now.getDate()).padStart(2, '0');
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const year = now.getFullYear();
  return `${day}.${month}.${year}`;
}

function hasRunValue(type: string | undefined, value: unknown): boolean {
  if (type === 'checkbox') return value === true;
  return String(value ?? '').trim().length > 0;
}

function runDefaultValueForField(field: { fieldName?: string; defaultValue?: string }): string {
  const configured = String(field?.defaultValue || '').trim();
  if (configured) return configured;
  const fieldName = String(field?.fieldName || '').trim();
  if (!fieldName) return '';
  if (RUN_DEFAULTS_BY_FIELD_NAME.has(fieldName)) {
    return String(RUN_DEFAULTS_BY_FIELD_NAME.get(fieldName) || '');
  }
  if (DATE_FIELD_NAME_PATTERN.test(fieldName)) {
    return todayDateLabel();
  }
  return '';
}

export function applyRunDefaultsFromModel(
  model: Record<string, unknown>,
  values: Record<string, unknown> = {}
): { values: Record<string, unknown>; changed: boolean } {
  const nextValues = { ...(values || {}) };
  let changed = false;

  const singleSections = (model?.single_sections as Array<{ fields?: Array<Record<string, unknown>> }>) || [];
  for (const section of singleSections) {
    for (const field of section?.fields || []) {
      if (field?.skipped === true) continue;
      const defaultValue = runDefaultValueForField(field as { fieldName?: string });
      if (!defaultValue) continue;
      const key = inputKeyForField(field as { fieldId: string });
      if (!key) continue;
      if (hasRunValue(String(field.type || 'text'), nextValues[key])) continue;
      nextValues[key] = defaultValue;
      changed = true;
    }
  }

  return { values: nextValues, changed };
}
