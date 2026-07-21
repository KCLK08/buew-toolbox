import { Asset } from 'expo-asset';
import { PDFDocument } from 'pdf-lib';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';

import { detectPdfFieldType } from './setup-model.js';

/** Bump when scan logic changes — triggers template re-scan on next load. */
export const ETB_SCAN_VERSION = 3;

type ScanField = {
  fieldName: string;
  labelCandidate: string;
  type: string;
  options: string[];
  page: number;
  orderIndex: number;
  rect: number[] | null;
  fieldId?: string;
};

type PdfAnnotation = {
  subtype?: string;
  fieldName?: string;
  rect?: number[];
  options?: unknown[];
};

type PdfTextItem = {
  str?: string;
  width?: number;
  transform?: number[];
};

type PdfJsPage = {
  getAnnotations: (options?: { intent: string }) => Promise<PdfAnnotation[]>;
  getTextContent: () => Promise<{ items: PdfTextItem[] }>;
};

type PdfJsDoc = {
  numPages: number;
  getPage: (pageNumber: number) => Promise<PdfJsPage>;
};

let pdfJsPromise: Promise<typeof pdfjs> | null = null;

async function loadPdfJs(): Promise<typeof pdfjs> {
  if (!pdfJsPromise) {
    pdfJsPromise = (async () => {
      const asset = Asset.fromModule(require('../../../../assets/pdf.worker.min.mjs'));
      await asset.downloadAsync();
      pdfjs.GlobalWorkerOptions.workerSrc = asset.localUri || asset.uri;
      return pdfjs;
    })();
  }
  return pdfJsPromise;
}

async function getPageAnnotations(page: PdfJsPage): Promise<PdfAnnotation[]> {
  try {
    return await page.getAnnotations({ intent: 'display' });
  } catch {
    return page.getAnnotations();
  }
}

function humanizeFieldName(value: string): string {
  return String(value || '')
    .replace(/[_\.]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeRect(rect: number[] | null | undefined) {
  if (!Array.isArray(rect) || rect.length < 4) {
    return null;
  }
  const x1 = Number(rect[0]);
  const y1 = Number(rect[1]);
  const x2 = Number(rect[2]);
  const y2 = Number(rect[3]);
  if (![x1, y1, x2, y2].every(Number.isFinite)) {
    return null;
  }
  const left = Math.min(x1, x2);
  const right = Math.max(x1, x2);
  const top = Math.max(y1, y2);
  const bottom = Math.min(y1, y2);
  return {
    left,
    right,
    top,
    bottom,
    width: Math.max(0, right - left),
    height: Math.max(0, top - bottom),
    centerX: (left + right) / 2,
    centerY: (top + bottom) / 2
  };
}

function uniqueStrings(values: string[] = []): string[] {
  return [...new Set(values.map((value) => String(value || '').trim()).filter(Boolean))];
}

function slugify(value: string): string {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function sortDetectedFields(fields: ScanField[]): ScanField[] {
  return [...fields].sort((left, right) => {
    if ((left.page ?? 9999) !== (right.page ?? 9999)) {
      return (left.page ?? 9999) - (right.page ?? 9999);
    }

    const leftRect = normalizeRect(left.rect);
    const rightRect = normalizeRect(right.rect);
    if (Boolean(leftRect) !== Boolean(rightRect)) {
      return leftRect ? -1 : 1;
    }

    if (leftRect && rightRect) {
      const topDelta = rightRect.top - leftRect.top;
      if (Math.abs(topDelta) > 4) {
        return topDelta;
      }
      const leftDelta = leftRect.left - rightRect.left;
      if (Math.abs(leftDelta) > 4) {
        return leftDelta;
      }
    }

    if ((left.orderIndex ?? 9999) !== (right.orderIndex ?? 9999)) {
      return (left.orderIndex ?? 9999) - (right.orderIndex ?? 9999);
    }

    return String(left.fieldName || '').localeCompare(String(right.fieldName || ''), 'de');
  });
}

function assignFieldIds(fields: ScanField[]): Array<ScanField & { fieldId: string }> {
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

function inferLabelCandidate(
  textLines: Array<{ text: string; x: number; y: number; width: number; centerX: number }>,
  rect: number[] | null,
  fallback = ''
): string {
  const normalizedRect = normalizeRect(rect);
  if (!normalizedRect) {
    return String(fallback || '').trim();
  }

  const fallbackLabel = String(fallback || '').trim() || 'Feld';
  const candidates: Array<{ text: string; score: number }> = [];

  for (const line of textLines) {
    const text = String(line.text || '').trim();
    if (!text || text.length < 2) {
      continue;
    }
    const left = Number(line.x ?? 0);
    const right = Number(line.x ?? 0) + Number(line.width ?? 0);
    const centerX = Number(line.centerX ?? left);
    const y = Number(line.y ?? 0);
    const isLeftLabel =
      right <= normalizedRect.left + 12 &&
      Math.abs(y - normalizedRect.centerY) <= Math.max(10, normalizedRect.height * 1.15);
    const isTopLabel =
      y > normalizedRect.top + 1 &&
      y <= normalizedRect.top + 45 &&
      centerX >= normalizedRect.left - 140 &&
      centerX <= normalizedRect.right + 140;

    if (!isLeftLabel && !isTopLabel) {
      continue;
    }

    const score = isLeftLabel
      ? Math.abs(normalizedRect.left - right) + Math.abs(y - normalizedRect.centerY) * 1.4
      : 20 + Math.abs(y - normalizedRect.top) * 1.3 + Math.abs(centerX - normalizedRect.centerX) * 0.55;

    candidates.push({ text, score });
  }

  candidates.sort((left, right) => left.score - right.score);
  const best = candidates[0]?.text;
  if (!best || best.length > 80) {
    return fallbackLabel;
  }
  return best;
}

function readSelectOptions(field: { getOptions?: () => string[] }, fallbackOptions: unknown[] = []): string[] {
  try {
    if (typeof field.getOptions === 'function') {
      return uniqueStrings(field.getOptions());
    }
  } catch {
    // Ignore fields that do not expose options directly.
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

async function extractWidgetMetadata(pdfJsDoc: PdfJsDoc) {
  const widgets = new Map<
    string,
    { page: number; orderIndex: number; rect: number[] | null; options: unknown[] }
  >();

  for (let page = 1; page <= pdfJsDoc.numPages; page += 1) {
    const annotations = await getPageAnnotations(await pdfJsDoc.getPage(page));
    annotations.forEach((annotation, index) => {
      if (annotation?.subtype !== 'Widget') {
        return;
      }
      const fieldName = String(annotation?.fieldName || '').trim();
      if (!fieldName) {
        return;
      }
      const candidate = {
        page,
        orderIndex: index,
        rect: Array.isArray(annotation.rect) ? annotation.rect.slice(0, 4) : null,
        options: annotation.options || []
      };
      const existing = widgets.get(fieldName);
      if (!existing) {
        widgets.set(fieldName, candidate);
        return;
      }
      if (candidate.page < existing.page || (candidate.page === existing.page && candidate.orderIndex < existing.orderIndex)) {
        widgets.set(fieldName, candidate);
      }
    });
  }
  return widgets;
}

async function extractTextLines(pdfJsDoc: PdfJsDoc) {
  const textLinesByPage = new Map<
    number,
    Array<{ text: string; x: number; y: number; width: number; centerX: number }>
  >();

  for (let page = 1; page <= pdfJsDoc.numPages; page += 1) {
    const textContent = await (await pdfJsDoc.getPage(page)).getTextContent();
    const textLines = [];
    for (const item of textContent.items || []) {
      const text = String(item?.str || '').trim();
      if (!text) {
        continue;
      }
      const transform = Array.isArray(item?.transform) ? item.transform : null;
      const x = Number(transform?.[4] ?? 0);
      const y = Number(transform?.[5] ?? 0);
      const width = Number(item?.width || 0);
      textLines.push({
        text,
        x,
        y,
        width,
        centerX: x + width / 2
      });
    }
    textLinesByPage.set(
      page,
      textLines.sort((left, right) => {
        if (Math.abs((right.y || 0) - (left.y || 0)) > 0.1) {
          return (right.y || 0) - (left.y || 0);
        }
        return (left.x || 0) - (right.x || 0);
      })
    );
  }
  return textLinesByPage;
}

export async function scanTemplatePdf(pdfBytes: Uint8Array) {
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

  const pdfJs = await loadPdfJs();
  const pdfJsDoc = (await pdfJs.getDocument({ data: pdfBytes }).promise) as PdfJsDoc;
  const widgetsByName = await extractWidgetMetadata(pdfJsDoc);
  const textLinesByPage = await extractTextLines(pdfJsDoc);

  const detectedFields = formFields.map((field, index) => {
    const fieldName = String(field.getName() || '').trim();
    const widget = widgetsByName.get(fieldName);
    const fieldType = detectPdfFieldType(field);
    const options = readSelectOptions(field as { getOptions?: () => string[] }, widget?.options || []);
    const page = Number(widget?.page || 1);
    const textLines = textLinesByPage.get(page) || [];
    const rect = Array.isArray(widget?.rect) ? widget.rect.slice(0, 4) : null;

    return {
      fieldName,
      labelCandidate: inferLabelCandidate(textLines, rect, humanizeFieldName(fieldName)),
      type: fieldType,
      options,
      page,
      orderIndex: Number(widget?.orderIndex ?? index),
      rect
    };
  });

  return {
    pageCount: Number(pdfJsDoc.numPages || pdfDoc.getPageCount() || 1),
    scanVersion: ETB_SCAN_VERSION,
    detectedFields: assignFieldIds(sortDetectedFields(detectedFields))
  };
}

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
