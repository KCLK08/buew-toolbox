import type { DetectedField } from '../types';
import {
  scanResultToTemplateFieldInput,
  type TemplateFieldInput
} from './template-field';
import type { ScanFieldResult } from './pdf-scan-types';

function fieldMatchKey(field: { fieldId: string; fieldName: string }): string {
  return String(field.fieldId || field.fieldName || '').trim();
}

/**
 * Merge freshly scanned AcroForm fields into existing template fields.
 * Manual fields are never removed or overwritten.
 */
export function mergeScannedFields(
  existing: DetectedField[],
  scanned: ScanFieldResult[]
): TemplateFieldInput[] {
  const manualFields = existing.filter((field) => field.source === 'manual');
  const manualKeys = new Set(manualFields.map((field) => fieldMatchKey(field)));

  const merged: TemplateFieldInput[] = manualFields.map((field) => ({
    fieldId: field.fieldId,
    fieldName: field.fieldName,
    labelCandidate: field.labelCandidate,
    type: field.type,
    options: field.options,
    page: field.page,
    orderIndex: field.orderIndex,
    source: field.source,
    geometry: field.geometry,
    rect: field.rect
  }));

  const acroById = new Map<string, TemplateFieldInput>();

  for (const scanField of scanned) {
    const key = fieldMatchKey(scanField);
    if (manualKeys.has(key)) continue;

    const prior = existing.find(
      (field) => field.source !== 'manual' && fieldMatchKey(field) === key
    );

    acroById.set(key, scanResultToTemplateFieldInput(scanField, 'acroform'));
    if (prior && prior.source === 'manual') {
      continue;
    }
  }

  for (const entry of acroById.values()) {
    merged.push(entry);
  }

  return merged.sort((left, right) => {
    const pageDelta = Number(left.page || 1) - Number(right.page || 1);
    if (pageDelta !== 0) return pageDelta;
    return Number(left.orderIndex || 0) - Number(right.orderIndex || 0);
  });
}

export function countAcroformDetected(scanned: ScanFieldResult[]): number {
  return scanned.filter(
    (field) => Array.isArray(field.rect) && field.rect.length >= 4
  ).length;
}
