import { Asset } from 'expo-asset';
import * as FileSystem from 'expo-file-system/legacy';

import type { PdfPreviewRuntimeAssets } from './pdf-preview-html';

const PDFJS_CORE_MODULE = require('../../../../assets/pdfjs/pdf.min.js');
const PDFJS_WORKER_MODULE = require('../../../../assets/pdfjs/pdf.worker.min.js');

let assetsPromise: Promise<PdfPreviewRuntimeAssets> | null = null;

/**
 * Loads bundled pdf.js 3.11.174 core source and resolves the worker URI via Expo Asset.
 * Results are cached for the app session — no CDN requests.
 */
export async function loadPdfPreviewAssets(): Promise<PdfPreviewRuntimeAssets> {
  if (!assetsPromise) {
    assetsPromise = loadPdfPreviewAssetsUncached();
  }
  return assetsPromise;
}

/** Clears the session cache — intended for tests only. */
export function resetPdfPreviewAssetsCache(): void {
  assetsPromise = null;
}

async function loadPdfPreviewAssetsUncached(): Promise<PdfPreviewRuntimeAssets> {
  const coreAsset = Asset.fromModule(PDFJS_CORE_MODULE);
  const workerAsset = Asset.fromModule(PDFJS_WORKER_MODULE);

  await Promise.all([coreAsset.downloadAsync(), workerAsset.downloadAsync()]);

  const coreUri = coreAsset.localUri || coreAsset.uri;
  const workerSrc = workerAsset.localUri || workerAsset.uri;

  if (!coreUri) {
    throw new Error('PDF preview core asset URI unavailable');
  }
  if (!workerSrc) {
    throw new Error('PDF preview worker asset URI unavailable');
  }

  const pdfJsSource = await FileSystem.readAsStringAsync(coreUri);

  return { pdfJsSource, workerSrc };
}
