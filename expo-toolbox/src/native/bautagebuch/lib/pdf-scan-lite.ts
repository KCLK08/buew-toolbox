import { PDFDocument, PDFDropdown, PDFRadioGroup } from 'pdf-lib';

import { detectPdfFieldType } from './setup-model.js';

/** Bump when scan logic changes — triggers template re-scan on next load. */
export const ETB_SCAN_VERSION = 2;

function humanizeFieldName(value: string): string {
  return String(value || '')
    .replace(/[_\.]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function slugify(value: string): string {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function uniqueStrings(values: string[] = []): string[] {
  return [...new Set(values.map((value) => String(value || '').trim()).filter(Boolean))];
}

function readSelectOptions(field: { getOptions?: () => string[] }, fallbackOptions: unknown[] = []): string[] {
  try {
    if (field instanceof PDFDropdown || field instanceof PDFRadioGroup) {
      return uniqueStrings(field.getOptions());
    }
    if (typeof field.getOptions === 'function') {
      return uniqueStrings(field.getOptions());
    }
  } catch {
    // Some fields do not expose options.
  }

  if (Array.isArray(fallbackOptions)) {
    return uniqueStrings(
      fallbackOptions.map((option) => {
        if (typeof option === 'string') return option;
        const record = option as { displayValue?: string; exportValue?: string; value?: string };
        return record.displayValue || record.exportValue || record.value || '';
      })
    );
  }
  return [];
}

function assignFieldIds(
  fields: Array<{
    fieldName: string;
    labelCandidate: string;
    type: string;
    options: string[];
    page: number;
    orderIndex: number;
    rect: number[] | null;
  }>
) {
  const usedIds = new Set<string>();
  return fields.map((field, index) => {
    const slug = slugify(field.fieldName) || `field-${index + 1}`;
    const page = Number(field.page || 1);
    const orderIndex = Number(field.orderIndex ?? index + 1);
    const baseId = `${slug}-p${page}-o${orderIndex}`;
    let fieldId = baseId;
    let suffix = 2;
    while (usedIds.has(fieldId)) {
      fieldId = `${baseId}-${suffix}`;
      suffix += 1;
    }
    usedIds.add(fieldId);
    return { ...field, fieldId };
  });
}

function sortDetectedFields<T extends { fieldName: string; page: number; orderIndex: number }>(fields: T[]): T[] {
  return [...fields].sort((left, right) => {
    if (left.page !== right.page) return left.page - right.page;
    if (left.orderIndex !== right.orderIndex) return left.orderIndex - right.orderIndex;
    return left.fieldName.localeCompare(right.fieldName, 'de');
  });
}

export async function scanTemplatePdfLite(pdfBytes: Uint8Array) {
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
    return {
      fieldName,
      labelCandidate: humanizeFieldName(fieldName),
      type: fieldType,
      options,
      page: 1,
      orderIndex: index,
      rect: null as number[] | null
    };
  });

  return {
    pageCount: pdfDoc.getPageCount(),
    scanVersion: ETB_SCAN_VERSION,
    detectedFields: assignFieldIds(sortDetectedFields(detectedFields))
  };
}

export function detectedFieldsNeedRescan(
  fields: Array<{ type?: string; options?: string[] }>
): boolean {
  return fields.some((field) => {
    const type = String(field.type || '').toLowerCase();
    if (type !== 'dropdown' && type !== 'radio') return false;
    return !Array.isArray(field.options) || field.options.length === 0;
  });
}
