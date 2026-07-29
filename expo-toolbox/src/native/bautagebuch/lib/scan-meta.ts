import type { FieldGeometry } from '../types';
import { fieldHasGeometry } from './template-field';

/** Bump when scan logic changes — triggers template re-scan on next load. */
export const ETB_SCAN_VERSION = 4;

type RescanField = {
  fieldId?: string;
  type?: string;
  options?: string[];
  rect?: number[] | null;
  geometry?: FieldGeometry | null;
  page?: number;
  source?: string;
};

export function detectedFieldsNeedRescan(fields: RescanField[]): boolean {
  if (fields.length === 0) {
    return false;
  }

  if (fields.some((field) => !String(field.fieldId || '').trim())) {
    return true;
  }

  if (fields.some((field) => String(field.type || '').toLowerCase() === 'unsupported')) {
    return true;
  }

  const needsGeometry = fields.filter((field) => {
    const source = String(field.source || 'acroform');
    return source !== 'manual';
  });
  if (needsGeometry.length > 0 && needsGeometry.every((field) => !fieldHasGeometry(field))) {
    return true;
  }

  return fields.some((field) => {
    const type = String(field.type || '').toLowerCase();
    if (type !== 'dropdown' && type !== 'radio') return false;
    return !Array.isArray(field.options) || field.options.length === 0;
  });
}
