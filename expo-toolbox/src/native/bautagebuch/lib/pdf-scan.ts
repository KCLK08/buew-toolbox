import { scanTemplatePdfLite } from './pdf-scan-lite';

export { ETB_SCAN_VERSION, detectedFieldsNeedRescan } from './scan-meta';
export type { PdfScanResult, ScanFieldResult } from './pdf-scan-types';

/**
 * Native scan: pdf-lib AcroForm extraction only.
 * pdfjs-dist is excluded from the Hermes bundle (dynamic import() unsupported).
 */
export async function scanTemplatePdf(pdfBytes: Uint8Array) {
  return scanTemplatePdfLite(pdfBytes);
}
