/** Raw field shape returned by scan pipelines (before DB persistence). */
export type ScanFieldResult = {
  fieldId: string;
  fieldName: string;
  labelCandidate: string;
  type: string;
  options: string[];
  page: number;
  orderIndex: number;
  rect: number[] | null;
};

/** Unified scan response for Web and Native. */
export type PdfScanResult = {
  fields: ScanFieldResult[];
  pageCount: number;
  scanVersion: string;
  hasRects: boolean;
};

/** @deprecated Use `fields` — kept for existing callers during migration. */
export type PdfScanResultLegacy = PdfScanResult & {
  detectedFields: ScanFieldResult[];
};

export function withDetectedFieldsAlias(result: PdfScanResult): PdfScanResultLegacy {
  return {
    ...result,
    detectedFields: result.fields
  };
}

export function resultHasRects(fields: ScanFieldResult[]): boolean {
  return fields.some((field) => Array.isArray(field.rect) && field.rect.length >= 4);
}
