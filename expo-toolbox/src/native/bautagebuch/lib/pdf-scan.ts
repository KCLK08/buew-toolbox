import { scanTemplatePdfFull } from './pdf-scan-full';
import { scanTemplatePdfLite } from './pdf-scan-lite';

export { ETB_SCAN_VERSION, detectedFieldsNeedRescan } from './scan-meta';
export type { PdfScanResult, ScanFieldResult } from './pdf-scan-types';

/**
 * Native scan: full pdf.js widget extraction first, lite AcroForm fallback.
 */
export async function scanTemplatePdf(pdfBytes: Uint8Array) {
  try {
    return await scanTemplatePdfFull(pdfBytes);
  } catch {
    return scanTemplatePdfLite(pdfBytes);
  }
}
