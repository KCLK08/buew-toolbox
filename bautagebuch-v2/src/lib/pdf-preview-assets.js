import pdfJsSource from 'pdfjs-dist/legacy/build/pdf.min.mjs?raw';
import workerSource from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?raw';

/** @typedef {{ pdfJsSource: string; workerSrc: string; workerSource: string }} PdfPreviewRuntimeAssets */

let assetsPromise = null;

export async function loadPdfPreviewAssets() {
  if (!assetsPromise) {
    assetsPromise = Promise.resolve({
      pdfJsSource,
      workerSource,
      workerSrc: 'inline://bautagebuch-pdf-preview-worker'
    });
  }
  return assetsPromise;
}

/** @param {import('./pdf-preview-html.js').PdfPreviewRuntimeAssets} assets */
export function resetPdfPreviewAssetsCache() {
  assetsPromise = null;
}
