import { PDFDocument } from 'pdf-lib';

import { detectPdfFieldType } from './setup-model.js';
import { ETB_SCAN_VERSION } from './scan-meta';
import {
  assignFieldIds,
  humanizeFieldName,
  inferLabelCandidate,
  readSelectOptions,
  sortDetectedFields,
  type MutableScanField
} from './pdf-scan-shared';
import { resultHasRects, withDetectedFieldsAlias, type PdfScanResultLegacy } from './pdf-scan-types';

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

type PdfJsModule = typeof import('pdfjs-dist/legacy/build/pdf.mjs');

let pdfJsPromise: Promise<PdfJsModule> | null = null;

async function resolvePdfJsWorkerSrc(): Promise<string> {
  try {
    const { Asset } = await import('expo-asset');
    const asset = Asset.fromModule(require('../../../../assets/pdf.worker.min.mjs'));
    await asset.downloadAsync();
    const uri = asset.localUri || asset.uri;
    if (uri) return uri;
  } catch {
    // Expo Asset unavailable (e.g. Node tests) — fall back below.
  }

  try {
    return require.resolve('pdfjs-dist/legacy/build/pdf.worker.min.mjs');
  } catch {
    throw new Error('pdf.js worker could not be resolved');
  }
}

async function loadPdfJs(): Promise<PdfJsModule> {
  if (!pdfJsPromise) {
    pdfJsPromise = (async () => {
      const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
      pdfjs.GlobalWorkerOptions.workerSrc = await resolvePdfJsWorkerSrc();
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

async function extractWidgetMetadata(pdfJsDoc: PdfJsDoc) {
  const widgets = new Map<
    string,
    { page: number; orderIndex: number; rect: number[] | null; options: unknown[] }
  >();

  for (let page = 1; page <= pdfJsDoc.numPages; page += 1) {
    const annotations = await getPageAnnotations(await pdfJsDoc.getPage(page));
    annotations.forEach((annotation, index) => {
      if (annotation?.subtype !== 'Widget') return;
      const fieldName = String(annotation?.fieldName || '').trim();
      if (!fieldName) return;

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
      if (
        candidate.page < existing.page ||
        (candidate.page === existing.page && candidate.orderIndex < existing.orderIndex)
      ) {
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
      if (!text) continue;
      const transform = Array.isArray(item?.transform) ? item.transform : null;
      const x = Number(transform?.[4] ?? 0);
      const y = Number(transform?.[5] ?? 0);
      const width = Number(item?.width || 0);
      textLines.push({ text, x, y, width, centerX: x + width / 2 });
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

function buildScanResult(pageCount: number, rawFields: MutableScanField[]): PdfScanResultLegacy {
  const fields = assignFieldIds(sortDetectedFields(rawFields));
  const result = {
    fields,
    pageCount,
    scanVersion: String(ETB_SCAN_VERSION),
    hasRects: resultHasRects(fields)
  };
  return withDetectedFieldsAlias(result);
}

export async function scanTemplatePdfFull(pdfBytes: Uint8Array): Promise<PdfScanResultLegacy> {
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

  return buildScanResult(Number(pdfJsDoc.numPages || pdfDoc.getPageCount() || 1), detectedFields);
}

/** Primary export name used by web platform file. */
export async function scanTemplatePdf(pdfBytes: Uint8Array): Promise<PdfScanResultLegacy> {
  return scanTemplatePdfFull(pdfBytes);
}

export { ETB_SCAN_VERSION } from './scan-meta';
export { detectedFieldsNeedRescan } from './scan-meta';
