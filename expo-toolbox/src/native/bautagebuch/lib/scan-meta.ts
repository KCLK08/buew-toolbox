/** Bump when scan logic changes — triggers template re-scan on next load. */
export const ETB_SCAN_VERSION = 3;

export function detectedFieldsNeedRescan(
  fields: Array<{ type?: string; options?: string[]; rect?: number[] | null }>
): boolean {
  if (fields.length === 0) {
    return true;
  }

  if (fields.some((field) => !Array.isArray(field.rect) || field.rect.length < 4)) {
    return true;
  }

  return fields.some((field) => {
    const type = String(field.type || '').toLowerCase();
    if (type !== 'dropdown' && type !== 'radio') return false;
    return !Array.isArray(field.options) || field.options.length === 0;
  });
}
