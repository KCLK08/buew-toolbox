export const PDF_PREVIEW_LOAD_ERROR = 'PDF Vorschau konnte nicht geladen werden';

/** @typedef {{ pdfJsSource: string; workerSrc: string; workerSource: string }} PdfPreviewRuntimeAssets */

/** @typedef {PdfPreviewRuntimeAssets & { base64: string }} BuildSimplePreviewHtmlOptions */

export function escapeInlineScript(source) {
  return source.replace(/<\/script/gi, '<\\/script');
}

export function escapeJsStringLiteral(value) {
  return String(value || '')
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

export function buildPdfJsInlineScript(pdfJsSource) {
  return `<script>\n${escapeInlineScript(pdfJsSource)}\n</script>`;
}

export function buildPdfWorkerInlineScript(workerSource) {
  return `<script>\n${escapeInlineScript(workerSource)}\n</script>`;
}

function previewErrorStyles() {
  return `
      #errorState {
        display: none;
        position: fixed;
        inset: 0;
        z-index: 100;
        align-items: center;
        justify-content: center;
        padding: 24px;
        background: rgba(242, 240, 235, 0.96);
        font: 15px/1.45 system-ui, sans-serif;
        color: #8b3a32;
        text-align: center;
      }
      #errorState.visible { display: flex; }`;
}

function previewBootHelpers(errorMessage, workerSrc) {
  const safeWorkerSrc = escapeJsStringLiteral(workerSrc);
  const safeErrorMessage = escapeJsStringLiteral(errorMessage);

  return `
      const PREVIEW_ERROR = '${safeErrorMessage}';

      function post(payload) {
        const serialized = JSON.stringify(payload);
        if (window.ReactNativeWebView && window.ReactNativeWebView.postMessage) {
          window.ReactNativeWebView.postMessage(serialized);
        }
        try {
          if (window.parent && window.parent !== window) {
            window.parent.postMessage(serialized, '*');
          }
        } catch (postError) {
          console.error('PDF preview parent post failed', postError);
        }
      }

      function showPreviewError(context, error) {
        console.error(context, error);
        const errorEl = document.getElementById('errorState');
        if (errorEl) errorEl.classList.add('visible');
        post({
          type: 'state',
          page: 1,
          pageCount: 1,
          ready: false,
          error: PREVIEW_ERROR
        });
      }

      const pdfjsLib = window.pdfjsLib;
      if (!pdfjsLib) {
        showPreviewError('PDF preview worker bootstrap failed', new Error('pdf.js core missing'));
      } else if (!globalThis.pdfjsWorker?.WorkerMessageHandler) {
        showPreviewError('PDF preview worker bootstrap failed', new Error('pdf.js worker missing'));
      } else {
        window.Worker = class PdfPreviewWorkerDisabled {
          constructor() {
            throw new Error('PDF preview: native Worker disabled');
          }
        };
        pdfjsLib.GlobalWorkerOptions.workerSrc = '${safeWorkerSrc}';
      }`;
}

export function buildSimplePdfPreviewHtml(options) {
  return buildScrollablePdfPreviewHtml(options);
}

/** Live BTB preview: all pages stacked, scroll + pinch zoom (no page picker). */
export function buildScrollablePdfPreviewHtml(options) {
  const { base64, pdfJsSource, workerSrc, workerSource } = options;

  return `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
    ${buildPdfWorkerInlineScript(workerSource)}
    ${buildPdfJsInlineScript(pdfJsSource)}
    <style>
      * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
      html, body {
        margin: 0; padding: 0; width: 100%; height: 100%;
        background: #f2f0eb; overflow: hidden;
      }
      #viewport {
        width: 100%; height: 100%; overflow: auto;
        -webkit-overflow-scrolling: touch;
        overscroll-behavior: contain;
      }
      #wrap {
        display: flex; flex-direction: column; align-items: center;
        gap: 12px; padding: 8px 0 16px;
        transform-origin: center top;
        will-change: transform;
      }
      .page-sheet {
        display: flex; justify-content: center; width: 100%;
      }
      canvas {
        display: block; max-width: 100%; height: auto;
        background: #fff; box-shadow: 0 1px 6px rgba(0,0,0,0.06);
      }
      #zoomHint {
        position: fixed; bottom: 8px; right: 8px; z-index: 50;
        font: 11px/1.3 system-ui, sans-serif; color: #666;
        background: rgba(255,255,255,0.88); padding: 4px 8px; border-radius: 8px;
        pointer-events: none; opacity: 0.85;
      }
      ${previewErrorStyles()}
    </style>
  </head>
  <body>
    <div id="errorState">${PDF_PREVIEW_LOAD_ERROR}</div>
    <div id="viewport">
      <div id="wrap"></div>
    </div>
    <div id="zoomHint">Scrollen · Zwei Finger zum Zoomen</div>
    <script>
      ${previewBootHelpers(PDF_PREVIEW_LOAD_ERROR, workerSrc)}

      let pdfDoc = null;
      let pinchScale = 1;
      let pinchStartDistance = 0;
      let pinchStartScale = 1;

      const MAX_CANVAS_PIXELS = 12000000;
      const viewport = document.getElementById('viewport');
      const wrap = document.getElementById('wrap');

      function resetPinchZoom() {
        pinchScale = 1;
        wrap.style.transform = 'scale(1)';
      }

      function touchDistance(touches) {
        const dx = touches[0].clientX - touches[1].clientX;
        const dy = touches[0].clientY - touches[1].clientY;
        return Math.hypot(dx, dy);
      }

      function applyPinchScale() {
        wrap.style.transform = 'scale(' + pinchScale + ')';
        wrap.style.transformOrigin = 'center top';
      }

      viewport.addEventListener('touchstart', (event) => {
        if (event.touches.length === 2) {
          pinchStartDistance = touchDistance(event.touches);
          pinchStartScale = pinchScale;
        }
      }, { passive: true });

      viewport.addEventListener('touchmove', (event) => {
        if (event.touches.length !== 2 || pinchStartDistance <= 0) return;
        event.preventDefault();
        const ratio = touchDistance(event.touches) / pinchStartDistance;
        pinchScale = Math.min(4, Math.max(1, pinchStartScale * ratio));
        applyPinchScale();
      }, { passive: false });

      viewport.addEventListener('touchend', () => {
        if (pinchScale <= 1.01) resetPinchZoom();
      }, { passive: true });

      async function renderPageSheet(pageNumber) {
        const page = await pdfDoc.getPage(pageNumber);
        const sheet = document.createElement('div');
        sheet.className = 'page-sheet';
        sheet.dataset.page = String(pageNumber);
        const canvas = document.createElement('canvas');
        sheet.appendChild(canvas);
        wrap.appendChild(sheet);

        const context = canvas.getContext('2d');
        const baseViewport = page.getViewport({ scale: 1 });
        const dpr = Math.min(window.devicePixelRatio || 1, 3);
        const maxCssWidth = window.innerWidth || 860;
        const fitScale = maxCssWidth / baseViewport.width;
        const viewportScale = Math.min(fitScale, 2.6);

        let pixelScale = viewportScale * dpr;
        const pixelArea = baseViewport.width * baseViewport.height * pixelScale * pixelScale;
        if (pixelArea > MAX_CANVAS_PIXELS) {
          pixelScale = Math.sqrt(MAX_CANVAS_PIXELS / (baseViewport.width * baseViewport.height));
        }

        const viewportObj = page.getViewport({ scale: pixelScale });
        canvas.width = Math.ceil(viewportObj.width);
        canvas.height = Math.ceil(viewportObj.height);
        canvas.style.width = Math.ceil(viewportObj.width / dpr) + 'px';
        canvas.style.height = Math.ceil(viewportObj.height / dpr) + 'px';

        await page.render({ canvasContext: context, viewport: viewportObj }).promise;
      }

      async function boot() {
        if (!pdfjsLib) return;
        const started = performance.now();
        try {
          const data = atob('${base64}');
          const bytes = new Uint8Array(data.length);
          for (let i = 0; i < data.length; i += 1) bytes[i] = data.charCodeAt(i);
          pdfDoc = await pdfjsLib.getDocument({ data: bytes }).promise;
          wrap.innerHTML = '';
          for (let pageNumber = 1; pageNumber <= pdfDoc.numPages; pageNumber += 1) {
            await renderPageSheet(pageNumber);
          }
          resetPinchZoom();
          post({
            type: 'state',
            page: 1,
            pageCount: pdfDoc.numPages,
            ready: true,
            error: null,
            renderMs: Math.round(performance.now() - started)
          });
        } catch (error) {
          showPreviewError('PDF preview boot failed', error);
        }
      }

      void boot();
    </script>
  </body>
</html>`;
}
