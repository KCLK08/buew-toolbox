import type { PDFDocument, PDFField } from 'pdf-lib';

import { normalizeRect } from './pdf-scan-shared';

export type AcroFormWidgetMetadata = {
  page: number;
  orderIndex: number;
  rect: number[] | null;
};

type PdfWidget = {
  getRectangle: () => { x: number; y: number; width: number; height: number };
  P?: () => unknown;
  dict: unknown;
};

function rectangleToLegacyRect(rect: { x: number; y: number; width: number; height: number }): number[] | null {
  const { x, y, width, height } = rect;
  if (![x, y, width, height].every(Number.isFinite)) {
    return null;
  }
  if (width <= 0 || height <= 0) {
    return null;
  }
  return [x, y, x + width, y + height];
}

function resolveWidgetPage(pdfDoc: PDFDocument, widget: PdfWidget): number | null {
  const pages = pdfDoc.getPages();
  const pageRef = typeof widget.P === 'function' ? widget.P() : undefined;
  if (pageRef) {
    const matched = pages.find((page) => page.ref === pageRef);
    if (matched) {
      const index = pages.indexOf(matched);
      return index >= 0 ? index + 1 : null;
    }
  }

  try {
    const context = (pdfDoc as { context?: { getObjectRef?: (object: unknown) => unknown } }).context;
    const widgetRef = context?.getObjectRef?.(widget.dict);
    if (widgetRef) {
      const page = pdfDoc.findPageForAnnotationRef(widgetRef as never);
      if (page) {
        const index = pages.indexOf(page);
        return index >= 0 ? index + 1 : null;
      }
    }
  } catch {
    // Some PDFs omit page refs on widgets — fall back below.
  }

  return null;
}

function pickPrimaryWidget(
  pdfDoc: PDFDocument,
  widgets: PdfWidget[],
  fallbackOrderIndex: number
): AcroFormWidgetMetadata {
  let best: AcroFormWidgetMetadata | null = null;

  widgets.forEach((widget, widgetIndex) => {
    const page = resolveWidgetPage(pdfDoc, widget) ?? 1;
    const rect = rectangleToLegacyRect(widget.getRectangle());
    const candidate: AcroFormWidgetMetadata = {
      page,
      orderIndex: widgetIndex,
      rect
    };

    if (!best) {
      best = candidate;
      return;
    }

    if (
      candidate.page < best.page ||
      (candidate.page === best.page && candidate.orderIndex < best.orderIndex)
    ) {
      best = candidate;
    }
  });

  return best ?? { page: 1, orderIndex: fallbackOrderIndex, rect: null };
}

export function extractFieldWidgetMetadata(
  pdfDoc: PDFDocument,
  field: PDFField,
  fallbackOrderIndex: number
): AcroFormWidgetMetadata {
  try {
    const widgets = field.acroField.getWidgets() as PdfWidget[];
    if (!Array.isArray(widgets) || widgets.length === 0) {
      return { page: 1, orderIndex: fallbackOrderIndex, rect: null };
    }
    return pickPrimaryWidget(pdfDoc, widgets, fallbackOrderIndex);
  } catch {
    return { page: 1, orderIndex: fallbackOrderIndex, rect: null };
  }
}

export function countScanFieldGeometry(fields: Array<{ rect?: number[] | null }>): {
  total: number;
  withGeometry: number;
  withoutGeometry: number;
} {
  const total = fields.length;
  const withGeometry = fields.filter((field) => normalizeRect(field.rect) !== null).length;
  return {
    total,
    withGeometry,
    withoutGeometry: Math.max(0, total - withGeometry)
  };
}

export function logAcroFormImportStats(fields: Array<{ rect?: number[] | null }>): void {
  const stats = countScanFieldGeometry(fields);
  console.log('[bautagebuch] AcroForm import');
  console.log(`  AcroForm fields: ${stats.total}`);
  console.log(`  Fields with geometry: ${stats.withGeometry}`);
  console.log(`  Fields without geometry: ${stats.withoutGeometry}`);
}
