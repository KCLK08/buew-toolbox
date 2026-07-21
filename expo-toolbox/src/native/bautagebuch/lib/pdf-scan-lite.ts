import { PDFDocument } from 'pdf-lib';

import { detectPdfFieldType } from '../lib/setup-model.js';

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

export async function scanTemplatePdfLite(pdfBytes: Uint8Array) {
  const pdfDoc = await PDFDocument.load(pdfBytes);
  const formFields = pdfDoc.getForm().getFields();
  if (!formFields.length) {
    throw new Error('Nur ausfüllbare AcroForm-PDFs werden unterstützt.');
  }

  const detectedFields = formFields.map((field, index) => {
    const fieldName = String(field.getName() || '').trim();
    const fieldType = detectPdfFieldType(field);
    return {
      fieldName,
      labelCandidate: humanizeFieldName(fieldName),
      type: fieldType,
      options: [],
      page: 1,
      orderIndex: index,
      rect: null as number[] | null
    };
  });

  return {
    pageCount: pdfDoc.getPageCount(),
    detectedFields: assignFieldIds(detectedFields)
  };
}
