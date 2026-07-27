/** Bump when scan logic changes — triggers template re-scan on next load. */
export const ETB_SCAN_VERSION = 4;

type RescanField = {
  fieldId?: string;
  type?: string;
  options?: string[];
  rect?: number[] | null;
  page?: number;
};

export function detectedFieldsNeedRescan(fields: RescanField[]): boolean {
  if (fields.length === 0) {
    return true;
  }

  if (fields.some((field) => !String(field.fieldId || '').trim())) {
    return true;
  }

  if (fields.some((field) => !Number(field.page || 0) || Number(field.page) < 1)) {
    return true;
  }

  if (fields.some((field) => !Array.isArray(field.rect) || field.rect.length < 4)) {
    return true;
  }

  if (fields.some((field) => String(field.type || '').toLowerCase() === 'unsupported')) {
    return true;
  }

  return fields.some((field) => {
    const type = String(field.type || '').toLowerCase();
    if (type !== 'dropdown' && type !== 'radio') return false;
    return !Array.isArray(field.options) || field.options.length === 0;
  });
}
