import { PDFDocument } from 'pdf-lib';

import { detectPdfFieldType } from './setup-model.js';
import { ETB_SCAN_VERSION } from './scan-meta';
import {
  extractFieldWidgetMetadata,
  logAcroFormImportStats
} from './pdf-scan-acroform-geometry';
import {
  assignFieldIds,
  humanizeFieldName,
  readSelectOptions,
  sortDetectedFields,
  type MutableScanField
} from './pdf-scan-shared';
import { resultHasRects, withDetectedFieldsAlias, type PdfScanResultLegacy } from './pdf-scan-types';

export { ETB_SCAN_VERSION, detectedFieldsNeedRescan } from './scan-meta';

function buildScanResult(pageCount: number, rawFields: MutableScanField[]): PdfScanResultLegacy {
  const fields = assignFieldIds(sortDetectedFields(rawFields));
  const result = {
    fields,
    pageCount,
    scanVersion: String(ETB_SCAN_VERSION),
    hasRects: resultHasRects(fields)
  };
  logAcroFormImportStats(fields);
  return withDetectedFieldsAlias(result);
}

export async function scanTemplatePdfLite(pdfBytes: Uint8Array): Promise<PdfScanResultLegacy> {
  let pdfDoc;
  try {
    pdfDoc = await PDFDocument.load(pdfBytes);
  } catch (error) {
    throw new Error(`PDF konnte nicht gelesen werden: ${error instanceof Error ? error.message : 'Unbekannt'}`);
  }

  let formFields;
  try {
    formFields = pdfDoc.getForm().getFields();
  } catch (error) {
    throw new Error(`AcroForm konnte nicht gelesen werden: ${error instanceof Error ? error.message : 'Unbekannt'}`);
  }

  if (!Array.isArray(formFields) || formFields.length === 0) {
    throw new Error('Nur ausfüllbare AcroForm-PDFs werden unterstützt.');
  }

  const detectedFields = formFields.map((field, index) => {
    const fieldName = String(field.getName() || '').trim();
    const fieldType = detectPdfFieldType(field);
    const options = readSelectOptions(field as { getOptions?: () => string[] });
    const widget = extractFieldWidgetMetadata(pdfDoc, field, index);

    return {
      fieldName,
      labelCandidate: humanizeFieldName(fieldName),
      type: fieldType,
      options,
      page: widget.page,
      orderIndex: widget.orderIndex,
      rect: widget.rect
    };
  });

  return buildScanResult(pdfDoc.getPageCount(), detectedFields);
}
