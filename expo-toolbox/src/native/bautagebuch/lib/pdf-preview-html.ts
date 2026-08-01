/** pdf.js 3.11.174 — bundled locally under assets/pdfjs/ (no CDN). */
export const PDFJS_VERSION = '3.11.174';

export const PDF_PREVIEW_LOAD_ERROR = 'PDF Vorschau konnte nicht geladen werden';

export type PreviewHtmlMode = 'default' | 'mapping' | 'assign' | 'overlay' | 'pinned';

export type PreviewHighlight = {
  fieldId: string;
  fieldName?: string;
  label?: string;
  source?: 'acroform' | 'manual' | 'ocr';
  index?: number;
  page: number;
  rect: number[];
};

export type PdfPreviewRuntimeAssets = {
  pdfJsSource: string;
  workerSrc: string;
  /** Worker script body — inlined in HTML head before core (file:// / blob workers fail in WebView). */
  workerSource: string;
};

export type BuildPreviewHtmlOptions = PdfPreviewRuntimeAssets & {
  base64: string;
  highlights?: PreviewHighlight[];
  assignedFieldIds?: string[];
  mode?: PreviewHtmlMode;
  highlightActive?: boolean;
  highQuality?: boolean;
};

export type BuildSimplePreviewHtmlOptions = PdfPreviewRuntimeAssets & {
  base64: string;
};

/** Prevents inline script content from prematurely closing the parent script tag. */
export function escapeInlineScript(source: string): string {
  return source.replace(/<\/script/gi, '<\\/script');
}

/** Escapes a value embedded in a single-quoted JavaScript string literal. */
export function escapeJsStringLiteral(value: string): string {
  return value
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

export function buildPdfJsInlineScript(pdfJsSource: string): string {
  return `<script>\n${escapeInlineScript(pdfJsSource)}\n</script>`;
}

export function buildPdfWorkerInlineScript(workerSource: string): string {
  return `<script>\n${escapeInlineScript(workerSource)}\n</script>`;
}

function previewErrorStyles(): string {
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

function previewBootHelpers(errorMessage: string, workerSrc: string): string {
  const safeWorkerSrc = escapeJsStringLiteral(workerSrc);
  const safeErrorMessage = escapeJsStringLiteral(errorMessage);

  return `
      const PREVIEW_ERROR = '${safeErrorMessage}';

      function post(payload) {
        if (window.ReactNativeWebView && window.ReactNativeWebView.postMessage) {
          window.ReactNativeWebView.postMessage(JSON.stringify(payload));
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

export function buildFieldPreviewHtml(options: BuildPreviewHtmlOptions): string {
  const {
    base64,
    pdfJsSource,
    workerSrc,
    workerSource,
    highlights = [],
    assignedFieldIds = [],
    mode = 'default',
    highlightActive = false,
    highQuality = false
  } = options;

  const mappingMode = mode === 'mapping' || mode === 'assign';
  const assignMode = mode === 'assign';
  const highlightsJson = JSON.stringify(highlights);
  const assignedFieldIdsJson = JSON.stringify(assignedFieldIds);

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
        position: relative; display: flex; justify-content: center;
        padding: ${mappingMode ? '4px 0 28px' : '8px 0'};
        transform-origin: 0 0;
        will-change: transform;
      }
      canvas {
        display: block; max-width: 100%; height: auto;
        background: #fff; box-shadow: 0 1px 6px rgba(0,0,0,0.06);
      }
      #overlay { position: absolute; inset: 0; pointer-events: none; }
      .highlight {
        position: absolute;
        border: 2px solid rgba(47, 111, 237, 0.45);
        background: rgba(47, 111, 237, 0.08);
        border-radius: 4px;
        transition: opacity 0.2s ease;
      }
      .highlight.selectable {
        pointer-events: auto;
        cursor: pointer;
      }
      .highlight.active {
        border: 3px solid rgba(196, 75, 50, 0.98);
        background: rgba(196, 75, 50, 0.28);
        box-shadow: 0 0 0 4px rgba(196, 75, 50, 0.3), 0 0 16px rgba(196, 75, 50, 0.35);
        animation: pulse 1s ease-in-out infinite;
        z-index: 3;
        opacity: 1;
      }
      .highlight.dim {
        opacity: 0.32;
        background: rgba(26, 25, 22, 0.12);
        border-color: rgba(26, 25, 22, 0.15);
        border-width: 1px;
      }
      .highlight.source-acroform:not(.active) {
        border-color: rgba(34, 139, 84, 0.55);
        background: rgba(34, 139, 84, 0.1);
      }
      .highlight.source-manual:not(.active) {
        border-color: rgba(196, 140, 40, 0.75);
        background: rgba(196, 140, 40, 0.12);
      }
      .highlight-label {
        position: absolute;
        left: 0;
        top: 0;
        max-width: 100%;
        padding: 2px 6px;
        font: 600 10px/1.3 system-ui, sans-serif;
        color: #fff;
        background: rgba(26, 25, 22, 0.78);
        border-radius: 4px 4px 4px 0;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        pointer-events: none;
      }
      .highlight-index {
        position: absolute;
        right: 4px;
        bottom: 2px;
        font: 700 10px/1 system-ui, sans-serif;
        color: rgba(26, 25, 22, 0.5);
        pointer-events: none;
      }
      .highlight.assigned {
        border: 2px solid rgba(46, 125, 50, 0.72);
        background: rgba(46, 125, 50, 0.08);
        opacity: 0.92;
      }
      .highlight.assigned.dim {
        opacity: 0.78;
        border-color: rgba(46, 125, 50, 0.55);
      }
      @keyframes pulse {
        0%, 100% { transform: scale(1); opacity: 1; }
        50% { transform: scale(1.03); opacity: 0.94; }
      }
      #zoomHint {
        position: fixed; bottom: 8px; right: 8px; z-index: 50;
        font: 11px/1.3 system-ui, sans-serif; color: #666;
        background: rgba(255,255,255,0.88); padding: 4px 8px; border-radius: 8px;
        pointer-events: none; opacity: 0; transition: opacity 0.3s;
      }
      #zoomHint.visible { opacity: 1; }
      ${previewErrorStyles()}
    </style>
  </head>
  <body>
    <div id="errorState">${PDF_PREVIEW_LOAD_ERROR}</div>
    <div id="viewport">
      <div id="wrap">
        <canvas id="canvas"></canvas>
        <div id="overlay"></div>
      </div>
    </div>
    <div id="zoomHint">Zwei Finger zum Zoomen</div>
    <script>
      const mappingMode = ${mappingMode ? 'true' : 'false'};
      const assignMode = ${assignMode ? 'true' : 'false'};
      const highlightActive = ${highlightActive ? 'true' : 'false'};
      const highQuality = ${highQuality ? 'true' : 'false'};
      ${previewBootHelpers(PDF_PREVIEW_LOAD_ERROR, workerSrc)}

      let highlights = ${highlightsJson};
      let assignedFieldIds = ${assignedFieldIdsJson};
      let pdfDoc = null;
      let currentPage = 1;
      let activeFieldId = '';
      let overlayPlacement = 'bottom';
      let viewportScale = 1;
      let renderTask = null;
      let pinchScale = 1;
      let pinchStartDistance = 0;
      let pinchStartScale = 1;

      const MAX_CANVAS_PIXELS = 12000000;
      const viewport = document.getElementById('viewport');
      const wrap = document.getElementById('wrap');
      const zoomHint = document.getElementById('zoomHint');

      function pageHighlights(pageNumber) {
        return highlights.filter((entry) => Number(entry.page || 1) === pageNumber);
      }

      function drawOverlay(pageNumber, viewportObj) {
        const overlay = document.getElementById('overlay');
        overlay.innerHTML = '';
        overlay.style.width = viewportObj.width + 'px';
        overlay.style.height = viewportObj.height + 'px';

        for (const entry of pageHighlights(pageNumber)) {
          if (!Array.isArray(entry.rect) || entry.rect.length < 4) continue;
          const [x1, y1, x2, y2] = entry.rect;
          const left = Math.min(x1, x2) * viewportScale;
          const top = (viewportObj.height / viewportScale - Math.max(y1, y2)) * viewportScale;
          const width = Math.abs(x2 - x1) * viewportScale;
          const height = Math.abs(y2 - y1) * viewportScale;
          const box = document.createElement('div');
          const isActive =
            String(entry.fieldId) === String(activeFieldId) ||
            String(entry.fieldName || '') === String(activeFieldId);
          const isAssigned = assignedFieldIds.includes(String(entry.fieldId));
          let className = 'highlight';
          if (isActive) className += ' active';
          else if (isAssigned) className += ' assigned';
          else if ((mappingMode || highlightActive) && activeFieldId) className += ' dim';
          box.className = className;
          box.style.left = left + 'px';
          box.style.top = top + 'px';
          box.style.width = width + 'px';
          box.style.height = height + 'px';
          if (entry.source === 'manual') box.classList.add('source-manual');
          else if (entry.source === 'acroform') box.classList.add('source-acroform');
          if (!assignMode) {
            const labelText = String(entry.label || entry.fieldName || '').trim();
            if (labelText) {
              const labelEl = document.createElement('div');
              labelEl.className = 'highlight-label';
              labelEl.textContent = labelText;
              box.appendChild(labelEl);
            }
            if (entry.index != null) {
              const indexEl = document.createElement('div');
              indexEl.className = 'highlight-index';
              indexEl.textContent = String(entry.index);
              box.appendChild(indexEl);
            }
          }
          if (assignMode) {
            box.classList.add('selectable');
            box.addEventListener('click', (event) => {
              event.stopPropagation();
              post({ type: 'fieldSelected', fieldId: String(entry.fieldId || '') });
            });
          }
          overlay.appendChild(box);
        }
      }

      function resetPinchZoom() {
        pinchScale = 1;
        wrap.style.transform = 'scale(1)';
        wrap.style.transformOrigin = '0 0';
      }

      function touchDistance(touches) {
        const dx = touches[0].clientX - touches[1].clientX;
        const dy = touches[0].clientY - touches[1].clientY;
        return Math.hypot(dx, dy);
      }

      function touchCenter(touches) {
        return {
          x: (touches[0].clientX + touches[1].clientX) / 2,
          y: (touches[0].clientY + touches[1].clientY) / 2
        };
      }

      function applyPinchScale() {
        wrap.style.transform = 'scale(' + pinchScale + ')';
        wrap.style.transformOrigin = '0 0';
      }

      viewport.addEventListener('touchstart', (event) => {
        if (event.touches.length === 2) {
          pinchStartDistance = touchDistance(event.touches);
          pinchStartScale = pinchScale;
          zoomHint.classList.add('visible');
        }
      }, { passive: true });

      viewport.addEventListener('touchmove', (event) => {
        if (event.touches.length !== 2 || pinchStartDistance <= 0) return;
        event.preventDefault();
        const newDistance = touchDistance(event.touches);
        const nextScale = Math.min(4, Math.max(1, pinchStartScale * (newDistance / pinchStartDistance)));
        const center = touchCenter(event.touches);
        const rect = viewport.getBoundingClientRect();
        const offsetX = center.x - rect.left;
        const offsetY = center.y - rect.top;
        const scaleRatio = nextScale / pinchScale;
        viewport.scrollLeft = (viewport.scrollLeft + offsetX) * scaleRatio - offsetX;
        viewport.scrollTop = (viewport.scrollTop + offsetY) * scaleRatio - offsetY;
        pinchScale = nextScale;
        applyPinchScale();
      }, { passive: false });

      viewport.addEventListener('touchend', (event) => {
        if (event.touches.length < 2) {
          pinchStartDistance = 0;
        }
        if (pinchScale <= 1.01) {
          resetPinchZoom();
        }
        setTimeout(() => zoomHint.classList.remove('visible'), 1200);
      }, { passive: true });

      function scrollActiveFieldErgonomic() {
        const active = document.querySelector('.highlight.active');
        if (!active || !viewport) return;

        const canvas = document.getElementById('canvas');
        const vpHeight = viewport.clientHeight;
        const fieldTop = active.offsetTop;
        const fieldHeight = active.offsetHeight || 24;
        const fieldCenter = fieldTop + fieldHeight / 2;

        let targetY;
        const upperThird = vpHeight * 0.28;
        const lowerThird = vpHeight * 0.62;
        const panelReserve = mappingMode ? vpHeight * 0.22 : 0;

        if (overlayPlacement === 'bottom') {
          targetY = fieldCenter - upperThird;
        } else if (overlayPlacement === 'top') {
          targetY = fieldCenter - lowerThird + panelReserve;
        } else if (overlayPlacement === 'left' || overlayPlacement === 'right') {
          targetY = fieldCenter - vpHeight * 0.38;
        } else {
          targetY = fieldCenter - vpHeight * 0.36;
        }

        const maxScroll = Math.max(0, (canvas ? canvas.offsetHeight : 0) + 40 - vpHeight);
        targetY = Math.min(Math.max(0, targetY), maxScroll);
        viewport.scrollTo({ top: targetY, behavior: 'smooth' });
      }

      async function renderPage(pageNumber) {
        if (!pdfDoc || !pdfjsLib) return;
        const started = performance.now();
        try {
          const safePage = Math.min(Math.max(pageNumber, 1), pdfDoc.numPages);
          currentPage = safePage;
          const page = await pdfDoc.getPage(safePage);
          const canvas = document.getElementById('canvas');
          const context = canvas.getContext('2d');
          const baseViewport = page.getViewport({ scale: 1 });
          const dpr = Math.min(window.devicePixelRatio || 1, highQuality ? 4 : 3);
          const maxCssWidth = Math.min(window.innerWidth || 860, mappingMode ? window.innerWidth : 860);
          const fitScale = maxCssWidth / baseViewport.width;
          const qualityBoost = highQuality ? 1.25 : 1;
          const scaleCap = mappingMode ? 3.6 : highQuality ? 3.4 : 2.6;
          viewportScale = Math.min(fitScale, scaleCap) * qualityBoost;

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

          if (renderTask) {
            try { renderTask.cancel(); } catch (_) {}
          }
          renderTask = page.render({ canvasContext: context, viewport: viewportObj });
          await renderTask.promise;
          renderTask = null;

          drawOverlay(safePage, viewportObj);
          resetPinchZoom();

          if (mappingMode && activeFieldId) {
            requestAnimationFrame(() => scrollActiveFieldErgonomic());
          }

          post({
            type: 'state',
            page: safePage,
            pageCount: pdfDoc.numPages,
            ready: true,
            error: null,
            renderMs: Math.round(performance.now() - started)
          });
        } catch (error) {
          showPreviewError('PDF preview render failed', error);
        }
      }

      async function boot() {
        if (!pdfjsLib) return;
        try {
          const data = atob('${base64}');
          const bytes = new Uint8Array(data.length);
          for (let i = 0; i < data.length; i += 1) bytes[i] = data.charCodeAt(i);
          pdfDoc = await pdfjsLib.getDocument({ data: bytes }).promise;
          await renderPage(1);
        } catch (error) {
          showPreviewError('PDF preview boot failed', error);
        }
      }

      function handleCommand(raw) {
        try {
          const message = typeof raw === 'string' ? JSON.parse(raw) : raw;
          if (!message || !message.type) return;
          if (message.type === 'setPage') {
            void renderPage(Number(message.page || 1));
          }
          if (message.type === 'setActive') {
            activeFieldId = String(message.fieldId || '');
            if (message.overlayPlacement) {
              overlayPlacement = String(message.overlayPlacement);
            }
            if (Array.isArray(message.assignedFieldIds)) {
              assignedFieldIds = message.assignedFieldIds.map((entry) => String(entry));
            }
            const page = Number(message.page || currentPage || 1);
            void renderPage(page);
          }
          if (message.type === 'resetZoom') {
            resetPinchZoom();
          }
          if (message.type === 'updateHighlights') {
            if (Array.isArray(message.highlights)) {
              highlights = message.highlights;
              void renderPage(Number(message.page || currentPage || 1));
            }
          }
        } catch (error) {
          console.error('PDF preview command failed', error);
        }
      }

      window.__applyPreviewCommand = handleCommand;
      document.addEventListener('message', (event) => handleCommand(event.data));
      window.addEventListener('message', (event) => handleCommand(event.data));
      void boot();
    </script>
  </body>
</html>`;
}

export function buildSimplePdfPreviewHtml(options: BuildSimplePreviewHtmlOptions | string): string {
  if (typeof options === 'string') {
    throw new Error('buildSimplePdfPreviewHtml requires bundled pdf.js assets');
  }

  return buildScrollablePdfPreviewHtml(options);
}

/** Live BTB preview: all pages stacked, scroll + pinch zoom (no page picker). */
export function buildScrollablePdfPreviewHtml(options: BuildSimplePreviewHtmlOptions): string {
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
      const zoomHint = document.getElementById('zoomHint');

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

      document.addEventListener('message', () => {});
      window.addEventListener('message', () => {});
      void boot();
    </script>
  </body>
</html>`;
}

/** Setup live preview: all pages stacked with field overlays, scroll + pinch zoom. */
export function buildScrollableFieldPreviewHtml(options: BuildPreviewHtmlOptions): string {
  const {
    base64,
    pdfJsSource,
    workerSrc,
    workerSource,
    highlights = [],
    highlightActive = false,
    highQuality = false
  } = options;

  const highlightsJson = JSON.stringify(highlights);

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
        transform-origin: 0 0;
        will-change: transform;
      }
      .page-sheet {
        display: flex; justify-content: center; width: 100%;
      }
      .page-frame {
        position: relative; display: inline-block;
      }
      canvas {
        display: block; max-width: 100%; height: auto;
        background: #fff; box-shadow: 0 1px 6px rgba(0,0,0,0.06);
      }
      .overlay {
        position: absolute; inset: 0; pointer-events: none;
      }
      .draw-layer {
        position: absolute; inset: 0; z-index: 4; pointer-events: none;
      }
      .page-frame.draw-enabled .draw-layer {
        pointer-events: auto; cursor: crosshair;
      }
      .draw-preview {
        position: absolute;
        border: 2px dashed rgba(196, 75, 50, 0.95);
        background: rgba(196, 75, 50, 0.15);
        border-radius: 4px;
        pointer-events: none;
      }
      .draft-rect {
        position: absolute;
        border: 3px solid rgba(196, 75, 50, 0.98);
        background: rgba(196, 75, 50, 0.22);
        border-radius: 4px;
        z-index: 6;
        box-shadow: 0 0 0 3px rgba(196, 75, 50, 0.22);
        pointer-events: none;
      }
      .draft-rect.editable {
        pointer-events: auto;
        touch-action: none;
        cursor: move;
      }
      .draft-drag-surface {
        position: absolute;
        inset: 0;
        cursor: move;
        touch-action: none;
      }
      .draft-handle {
        position: absolute;
        width: 16px;
        height: 16px;
        margin: -8px 0 0 -8px;
        border-radius: 999px;
        background: #fff;
        border: 2px solid rgba(196, 75, 50, 0.98);
        box-shadow: 0 1px 4px rgba(0,0,0,0.18);
        pointer-events: auto;
        touch-action: none;
        z-index: 2;
      }
      .draft-handle.side-n,
      .draft-handle.side-s {
        margin-left: -8px;
        left: 50%;
      }
      .draft-handle.side-e,
      .draft-handle.side-w {
        margin-top: -8px;
        top: 50%;
      }
      .highlight {
        position: absolute;
        border: 2px solid rgba(47, 111, 237, 0.45);
        background: rgba(47, 111, 237, 0.08);
        border-radius: 4px;
        transition: opacity 0.2s ease;
      }
      .highlight.selectable {
        pointer-events: auto;
        cursor: pointer;
      }
      .highlight.active {
        border: 3px solid rgba(196, 75, 50, 0.98);
        background: rgba(196, 75, 50, 0.28);
        box-shadow: 0 0 0 4px rgba(196, 75, 50, 0.3), 0 0 16px rgba(196, 75, 50, 0.35);
        animation: pulse 1s ease-in-out infinite;
        z-index: 3;
        opacity: 1;
      }
      .highlight.dim {
        opacity: 0.32;
        background: rgba(26, 25, 22, 0.12);
        border-color: rgba(26, 25, 22, 0.15);
        border-width: 1px;
      }
      .highlight.source-acroform:not(.active) {
        border-color: rgba(34, 139, 84, 0.55);
        background: rgba(34, 139, 84, 0.1);
      }
      .highlight.source-manual:not(.active) {
        border-color: rgba(196, 140, 40, 0.75);
        background: rgba(196, 140, 40, 0.12);
      }
      .highlight-label {
        position: absolute;
        left: 0;
        top: 0;
        max-width: 100%;
        padding: 2px 6px;
        font: 600 10px/1.3 system-ui, sans-serif;
        color: #fff;
        background: rgba(26, 25, 22, 0.78);
        border-radius: 4px 4px 4px 0;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        pointer-events: none;
      }
      .highlight-index {
        position: absolute;
        right: 4px;
        bottom: 2px;
        font: 700 10px/1 system-ui, sans-serif;
        color: rgba(26, 25, 22, 0.5);
        pointer-events: none;
      }
      @keyframes pulse {
        0%, 100% { transform: scale(1); opacity: 1; }
        50% { transform: scale(1.03); opacity: 0.94; }
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
      const highlightActive = ${highlightActive ? 'true' : 'false'};
      const highQuality = ${highQuality ? 'true' : 'false'};
      const framesOnly = true;
      ${previewBootHelpers(PDF_PREVIEW_LOAD_ERROR, workerSrc)}

      let highlights = ${highlightsJson};
      let pdfDoc = null;
      let activeFieldId = '';
      let pinchScale = 1;
      let pinchStartDistance = 0;
      let pinchStartScale = 1;
      const pageMeta = new Map();
      let drawModeEnabled = false;
      let drawStart = null;
      let drawPreviewEl = null;
      let drawPageNumber = 1;
      let draftRectState = null;
      let draftRectEl = null;
      let draftInteraction = null;
      let draftBottomReservePx = 0;

      const MAX_CANVAS_PIXELS = 12000000;
      const viewport = document.getElementById('viewport');
      const wrap = document.getElementById('wrap');

      function pageHighlights(pageNumber) {
        return highlights.filter((entry) => Number(entry.page || 1) === pageNumber);
      }

      function matchesActive(entry) {
        const active = String(activeFieldId || '');
        if (!active) return false;
        return String(entry.fieldId) === active || String(entry.fieldName || '') === active;
      }

      function drawPageOverlay(pageNumber, overlayEl, meta) {
        overlayEl.innerHTML = '';
        overlayEl.style.width = Math.ceil(meta.cssWidth) + 'px';
        overlayEl.style.height = Math.ceil(meta.cssHeight) + 'px';
        const scale = meta.cssScale;
        const pdfHeight = meta.baseViewportHeight;

        for (const entry of pageHighlights(pageNumber)) {
          if (!Array.isArray(entry.rect) || entry.rect.length < 4) continue;
          const [x1, y1, x2, y2] = entry.rect;
          const left = Math.min(x1, x2) * scale;
          const top = (pdfHeight - Math.max(y1, y2)) * scale;
          const width = Math.abs(x2 - x1) * scale;
          const height = Math.abs(y2 - y1) * scale;
          const box = document.createElement('div');
          const isActive = matchesActive(entry);
          let className = 'highlight';
          if (isActive) className += ' active';
          else if (highlightActive && activeFieldId) className += ' dim';
          box.className = className;
          box.style.left = left + 'px';
          box.style.top = top + 'px';
          box.style.width = width + 'px';
          box.style.height = height + 'px';
          if (entry.source === 'manual') box.classList.add('source-manual');
          else if (entry.source === 'acroform') box.classList.add('source-acroform');
          if (!framesOnly) {
            const labelText = String(entry.label || entry.fieldName || '').trim();
            if (labelText) {
              const labelEl = document.createElement('div');
              labelEl.className = 'highlight-label';
              labelEl.textContent = labelText;
              box.appendChild(labelEl);
            }
            if (entry.index != null) {
              const indexEl = document.createElement('div');
              indexEl.className = 'highlight-index';
              indexEl.textContent = String(entry.index);
              box.appendChild(indexEl);
            }
          }
          if (!drawModeEnabled) {
            box.classList.add('selectable');
            box.addEventListener('click', (event) => {
              event.stopPropagation();
              post({ type: 'fieldSelected', fieldId: String(entry.fieldId || '') });
            });
          }
          overlayEl.appendChild(box);
        }
      }

      function setDrawMode(enabled) {
        drawModeEnabled = Boolean(enabled);
        for (const frame of wrap.querySelectorAll('.page-frame')) {
          frame.classList.toggle('draw-enabled', drawModeEnabled);
        }
        if (drawModeEnabled) {
          viewport.style.overflow = 'hidden';
          viewport.style.touchAction = 'none';
          pinchStartDistance = 0;
        } else {
          restoreViewportNavigation();
        }
        if (!drawModeEnabled && drawPreviewEl) {
          drawPreviewEl.remove();
          drawPreviewEl = null;
          drawStart = null;
        }
      }

      function cssToPdfRect(pageNumber, cssX, cssY, cssWidth, cssHeight) {
        const meta = pageMeta.get(pageNumber);
        if (!meta) return null;
        const scale = meta.cssScale;
        const pdfHeight = meta.baseViewportHeight;
        const x = cssX / scale;
        const width = cssWidth / scale;
        const height = cssHeight / scale;
        const y = pdfHeight - (cssY + cssHeight) / scale;
        return { x, y, width, height };
      }

      function pdfToCssRect(pageNumber, pdfRect) {
        const meta = pageMeta.get(pageNumber);
        if (!meta || !pdfRect) return null;
        const scale = meta.cssScale;
        const pdfHeight = meta.baseViewportHeight;
        const left = Number(pdfRect.x || 0) * scale;
        const width = Number(pdfRect.width || 0) * scale;
        const height = Number(pdfRect.height || 0) * scale;
        const top = (pdfHeight - Number(pdfRect.y || 0) - Number(pdfRect.height || 0)) * scale;
        return { left, top, width, height };
      }

      function emitDraftRectUpdate() {
        if (!draftRectState) return;
        post({
          type: 'fieldDraftUpdated',
          page: draftRectState.page,
          rect: draftRectState.pdfRect
        });
      }

      function removeDraftRectEl() {
        if (draftRectEl) {
          draftRectEl.remove();
          draftRectEl = null;
        }
        draftInteraction = null;
      }

      function findPageFrame(pageNumber) {
        const sheet = wrap.querySelector('.page-sheet[data-page="' + pageNumber + '"]');
        return sheet ? sheet.querySelector('.page-frame') : null;
      }

      function applyDraftCssRect(cssRect) {
        if (!draftRectEl || !cssRect) return;
        draftRectEl.style.left = cssRect.left + 'px';
        draftRectEl.style.top = cssRect.top + 'px';
        draftRectEl.style.width = Math.max(8, cssRect.width) + 'px';
        draftRectEl.style.height = Math.max(8, cssRect.height) + 'px';
      }

      function applyDraftViewportMode() {
        if (!draftRectState) {
          restoreViewportNavigation();
          return;
        }
        if (draftRectState.editable) {
          viewport.style.overflow = 'auto';
          viewport.style.touchAction = 'pan-y pinch-zoom';
          return;
        }
        restoreViewportNavigation();
      }

      function restoreViewportNavigation() {
        if (drawModeEnabled) return;
        viewport.style.overflow = 'auto';
        viewport.style.touchAction = 'pan-y pinch-zoom';
      }

      function mountDraftHandles(frame, pageNumber) {
        if (!draftRectEl || !draftRectState?.editable) return;
        draftRectEl.querySelectorAll('.draft-handle, .draft-drag-surface').forEach((node) => node.remove());
        const dragSurface = document.createElement('div');
        dragSurface.className = 'draft-drag-surface';
        draftRectEl.appendChild(dragSurface);
        ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].forEach((handleName) => {
          const handle = document.createElement('div');
          handle.className = 'draft-handle';
          if (handleName.length === 1) handle.classList.add('side-' + handleName);
          handle.dataset.handle = handleName;
          if (handleName.includes('n')) handle.style.top = '0';
          if (handleName.includes('s')) handle.style.top = '100%';
          if (handleName.includes('w')) handle.style.left = '0';
          if (handleName.includes('e')) handle.style.left = '100%';
          if (handleName === 'n' || handleName === 's') handle.style.left = '50%';
          if (handleName === 'e' || handleName === 'w') handle.style.top = '50%';
          draftRectEl.appendChild(handle);
        });
        if (draftRectEl.dataset.interactions !== '1') {
          draftRectEl.dataset.interactions = '1';
          attachDraftRectInteractions(frame, pageNumber, dragSurface);
        }
      }

      function attachDraftRectInteractions(frame, pageNumber, dragSurface) {
        if (!draftRectEl || !draftRectState || !draftRectState.editable) return;

        const handles = draftRectEl.querySelectorAll('.draft-handle');
        const touchOpts = { passive: false };

        const beginInteraction = (clientX, clientY, mode, handleName) => {
          const cssRect = pdfToCssRect(pageNumber, draftRectState.pdfRect);
          if (!cssRect) return;
          draftInteraction = {
            mode,
            handle: handleName,
            startX: clientX,
            startY: clientY,
            origin: { ...cssRect }
          };
        };

        const updateInteraction = (clientX, clientY) => {
          if (!draftInteraction || !draftRectState) return;
          const dx = clientX - draftInteraction.startX;
          const dy = clientY - draftInteraction.startY;
          const origin = draftInteraction.origin;
          let next = { ...origin };

          if (draftInteraction.mode === 'move') {
            next.left = origin.left + dx;
            next.top = origin.top + dy;
          } else {
            const handle = draftInteraction.handle || 'se';
            if (handle.includes('e')) next.width = Math.max(8, origin.width + dx);
            if (handle.includes('s')) next.height = Math.max(8, origin.height + dy);
            if (handle.includes('w')) {
              next.width = Math.max(8, origin.width - dx);
              next.left = origin.left + dx;
            }
            if (handle.includes('n')) {
              next.height = Math.max(8, origin.height - dy);
              next.top = origin.top + dy;
            }
          }

          const maxW = frame.clientWidth;
          const maxH = frame.clientHeight;
          next.left = Math.max(0, Math.min(next.left, maxW - next.width));
          next.top = Math.max(0, Math.min(next.top, maxH - next.height));
          applyDraftCssRect(next);
        };

        const finishInteraction = () => {
          if (!draftInteraction || !draftRectState || !draftRectEl) return;
          const cssLeft = parseFloat(draftRectEl.style.left || '0');
          const cssTop = parseFloat(draftRectEl.style.top || '0');
          const cssWidth = parseFloat(draftRectEl.style.width || '0');
          const cssHeight = parseFloat(draftRectEl.style.height || '0');
          const pdfRect = cssToPdfRect(pageNumber, cssLeft, cssTop, cssWidth, cssHeight);
          draftInteraction = null;
          if (!pdfRect) return;
          draftRectState.pdfRect = pdfRect;
          emitDraftRectUpdate();
        };

        const bindTouchTarget = (target, mode, handleName) => {
          target.addEventListener('touchstart', (event) => {
            if (!draftRectState.editable || event.touches.length !== 1) return;
            event.preventDefault();
            event.stopPropagation();
            beginInteraction(event.touches[0].clientX, event.touches[0].clientY, mode, handleName);
          }, touchOpts);

          target.addEventListener('touchmove', (event) => {
            if (!draftInteraction || event.touches.length !== 1) return;
            event.preventDefault();
            event.stopPropagation();
            updateInteraction(event.touches[0].clientX, event.touches[0].clientY);
          }, touchOpts);

          target.addEventListener('touchend', (event) => {
            if (!draftInteraction) return;
            event.preventDefault();
            finishInteraction();
          }, touchOpts);

          target.addEventListener('touchcancel', () => {
            draftInteraction = null;
          });
        };

        bindTouchTarget(dragSurface, 'move');
        handles.forEach((handle) => {
          bindTouchTarget(handle, 'resize', String(handle.dataset.handle || 'se'));
        });

        dragSurface.addEventListener('pointerdown', (event) => {
          if (!draftRectState.editable || event.pointerType === 'touch') return;
          event.preventDefault();
          event.stopPropagation();
          beginInteraction(event.clientX, event.clientY, 'move');
          try { dragSurface.setPointerCapture(event.pointerId); } catch (_) {}
        });

        handles.forEach((handle) => {
          handle.addEventListener('pointerdown', (event) => {
            if (event.pointerType === 'touch') return;
            event.preventDefault();
            event.stopPropagation();
            beginInteraction(event.clientX, event.clientY, 'resize', String(handle.dataset.handle || 'se'));
            try { handle.setPointerCapture(event.pointerId); } catch (_) {}
          });
        });

        dragSurface.addEventListener('pointermove', (event) => {
          if (!draftInteraction || event.pointerType === 'touch') return;
          event.preventDefault();
          updateInteraction(event.clientX, event.clientY);
        });

        const onPointerEnd = (event) => {
          if (!draftInteraction || event.pointerType === 'touch') return;
          event.preventDefault();
          finishInteraction();
          try { dragSurface.releasePointerCapture(event.pointerId); } catch (_) {}
        };

        dragSurface.addEventListener('pointerup', onPointerEnd);
        dragSurface.addEventListener('pointercancel', onPointerEnd);
      }

      function renderDraftRect() {
        removeDraftRectEl();
        if (!draftRectState) return;
        const frame = findPageFrame(draftRectState.page);
        if (!frame) return;
        const layer = frame.querySelector('.draw-layer');
        if (!layer) return;
        const cssRect = pdfToCssRect(draftRectState.page, draftRectState.pdfRect);
        if (!cssRect) return;

        draftRectEl = document.createElement('div');
        draftRectEl.className = 'draft-rect' + (draftRectState.editable ? ' editable' : '');
        applyDraftCssRect(cssRect);

        layer.appendChild(draftRectEl);
        if (draftRectState.editable) {
          mountDraftHandles(frame, draftRectState.page);
        }
        applyDraftViewportMode();
        requestAnimationFrame(() => scrollToDraftRect());
      }

      function setDraftRect(message) {
        const page = Number(message.page || 1);
        const editable = Boolean(message.editable);
        const pdfRect = message.rect || null;

        if (
          draftRectState &&
          draftRectEl &&
          draftRectState.page === page &&
          draftRectState.editable === editable
        ) {
          draftRectState.pdfRect = pdfRect;
          const cssRect = pdfToCssRect(page, pdfRect);
          applyDraftCssRect(cssRect);
          return;
        }

        draftRectState = { page, pdfRect, editable };
        renderDraftRect();
      }

      function setDraftRectEditable(enabled, bottomReserve) {
        if (!draftRectState) return;
        if (Number.isFinite(Number(bottomReserve))) {
          draftBottomReservePx = Math.max(0, Number(bottomReserve));
        }
        const nextEditable = Boolean(enabled);
        if (draftRectState.editable === nextEditable && draftRectEl) {
          requestAnimationFrame(() => scrollToDraftRect());
          return;
        }

        const scrollTop = viewport.scrollTop;
        const scrollLeft = viewport.scrollLeft;
        draftRectState.editable = nextEditable;

        if (!draftRectEl) {
          renderDraftRect();
          return;
        }

        draftRectEl.classList.toggle('editable', nextEditable);
        const frame = findPageFrame(draftRectState.page);
        if (nextEditable && frame) {
          mountDraftHandles(frame, draftRectState.page);
        } else {
          draftRectEl.querySelectorAll('.draft-handle').forEach((handle) => handle.remove());
          draftInteraction = null;
        }
        applyDraftViewportMode();
        viewport.scrollTop = scrollTop;
        viewport.scrollLeft = scrollLeft;
        requestAnimationFrame(() => scrollToDraftRect());
      }

      function clearDraftRect() {
        draftRectState = null;
        removeDraftRectEl();
        restoreViewportNavigation();
      }

      function attachDrawHandlers(frame, pageNumber) {
        const layer = frame.querySelector('.draw-layer');
        if (!layer || layer.dataset.bound === '1') return;
        layer.dataset.bound = '1';

        const clearPreview = () => {
          if (drawPreviewEl) {
            drawPreviewEl.remove();
            drawPreviewEl = null;
          }
          drawStart = null;
        };

        const localPoint = (clientX, clientY) => {
          const rect = frame.getBoundingClientRect();
          return {
            x: clientX - rect.left,
            y: clientY - rect.top
          };
        };

        const updatePreview = (clientX, clientY) => {
          if (!drawStart || !drawPreviewEl) return;
          const current = localPoint(clientX, clientY);
          const left = Math.min(drawStart.x, current.x);
          const top = Math.min(drawStart.y, current.y);
          drawPreviewEl.style.left = left + 'px';
          drawPreviewEl.style.top = top + 'px';
          drawPreviewEl.style.width = Math.abs(current.x - drawStart.x) + 'px';
          drawPreviewEl.style.height = Math.abs(current.y - drawStart.y) + 'px';
        };

        const finishDraw = (clientX, clientY) => {
          if (!drawStart || !drawPreviewEl) return;
          const current = localPoint(clientX, clientY);
          const left = Math.min(drawStart.x, current.x);
          const top = Math.min(drawStart.y, current.y);
          const width = Math.abs(current.x - drawStart.x);
          const height = Math.abs(current.y - drawStart.y);
          clearPreview();
          if (width < 8 || height < 8) return;
          const pdfRect = cssToPdfRect(pageNumber, left, top, width, height);
          if (!pdfRect) return;
          drawModeEnabled = false;
          for (const frame of wrap.querySelectorAll('.page-frame')) {
            frame.classList.remove('draw-enabled');
          }
          post({
            type: 'fieldDrawDraft',
            page: pageNumber,
            rect: pdfRect
          });
          requestAnimationFrame(() => restoreViewportNavigation());
        };

        const beginDraw = (clientX, clientY) => {
          if (!drawModeEnabled) return;
          drawPageNumber = pageNumber;
          drawStart = localPoint(clientX, clientY);
          drawPreviewEl = document.createElement('div');
          drawPreviewEl.className = 'draw-preview';
          layer.appendChild(drawPreviewEl);
        };

        layer.addEventListener('touchstart', (event) => {
          if (!drawModeEnabled || event.touches.length !== 1) return;
          event.preventDefault();
          event.stopPropagation();
          beginDraw(event.touches[0].clientX, event.touches[0].clientY);
        }, { passive: false });

        layer.addEventListener('touchmove', (event) => {
          if (!drawModeEnabled || !drawStart || event.touches.length !== 1) return;
          event.preventDefault();
          event.stopPropagation();
          updatePreview(event.touches[0].clientX, event.touches[0].clientY);
        }, { passive: false });

        layer.addEventListener('touchend', (event) => {
          if (!drawModeEnabled || !drawStart) return;
          event.preventDefault();
          const touch = event.changedTouches[0];
          if (touch) finishDraw(touch.clientX, touch.clientY);
        }, { passive: false });

        layer.addEventListener('touchcancel', () => clearPreview());

        layer.addEventListener('pointerdown', (event) => {
          if (!drawModeEnabled || event.pointerType === 'touch') return;
          event.preventDefault();
          beginDraw(event.clientX, event.clientY);
          try { layer.setPointerCapture(event.pointerId); } catch (_) {}
        });

        layer.addEventListener('pointermove', (event) => {
          if (!drawModeEnabled || !drawStart || event.pointerType === 'touch') return;
          updatePreview(event.clientX, event.clientY);
        });

        layer.addEventListener('pointerup', (event) => {
          if (!drawModeEnabled || event.pointerType === 'touch') return;
          finishDraw(event.clientX, event.clientY);
          try { layer.releasePointerCapture(event.pointerId); } catch (_) {}
        });

        layer.addEventListener('pointercancel', () => clearPreview());
      }

      function refreshAllOverlays() {
        for (const sheet of wrap.querySelectorAll('.page-sheet')) {
          const pageNumber = Number(sheet.dataset.page || 1);
          const meta = pageMeta.get(pageNumber);
          const overlayEl = sheet.querySelector('.overlay');
          if (!meta || !overlayEl) continue;
          drawPageOverlay(pageNumber, overlayEl, meta);
        }
      }

      function scrollToActiveField() {
        const active = document.querySelector('.highlight.active');
        if (!active || !viewport) return;
        const sheet = active.closest('.page-sheet');
        if (!sheet) return;
        const sheetTop = sheet.offsetTop;
        const fieldTop = active.offsetTop;
        const fieldHeight = active.offsetHeight || 24;
        const targetY = sheetTop + fieldTop - viewport.clientHeight * 0.32 + fieldHeight / 2;
        viewport.scrollTo({ top: Math.max(0, targetY), behavior: 'smooth' });
      }

      function scrollToDraftRect() {
        if (!draftRectEl || !viewport) return;
        const viewportRect = viewport.getBoundingClientRect();
        const elRect = draftRectEl.getBoundingClientRect();
        const topPad = 20;
        const bottomPad = draftBottomReservePx + 20;
        const visibleTop = viewportRect.top + topPad;
        const visibleBottom = viewportRect.bottom - bottomPad;

        if (elRect.top < visibleTop) {
          viewport.scrollTop += elRect.top - visibleTop;
        } else if (elRect.bottom > visibleBottom) {
          viewport.scrollTop += elRect.bottom - visibleBottom;
        }
      }

      function resetPinchZoom() {
        pinchScale = 1;
        wrap.style.transform = 'scale(1)';
        wrap.style.transformOrigin = '0 0';
      }

      function touchDistance(touches) {
        const dx = touches[0].clientX - touches[1].clientX;
        const dy = touches[0].clientY - touches[1].clientY;
        return Math.hypot(dx, dy);
      }

      function touchCenter(touches) {
        return {
          x: (touches[0].clientX + touches[1].clientX) / 2,
          y: (touches[0].clientY + touches[1].clientY) / 2
        };
      }

      function applyPinchScale() {
        wrap.style.transform = 'scale(' + pinchScale + ')';
        wrap.style.transformOrigin = '0 0';
      }

      viewport.addEventListener('touchstart', (event) => {
        if (drawModeEnabled) return;
        if (event.touches.length === 2) {
          pinchStartDistance = touchDistance(event.touches);
          pinchStartScale = pinchScale;
        }
      }, { passive: true });

      viewport.addEventListener('touchmove', (event) => {
        if (drawModeEnabled) return;
        if (event.touches.length !== 2 || pinchStartDistance <= 0) return;
        event.preventDefault();
        const newDistance = touchDistance(event.touches);
        const nextScale = Math.min(4, Math.max(1, pinchStartScale * (newDistance / pinchStartDistance)));
        const center = touchCenter(event.touches);
        const rect = viewport.getBoundingClientRect();
        const offsetX = center.x - rect.left;
        const offsetY = center.y - rect.top;
        const scaleRatio = nextScale / pinchScale;
        viewport.scrollLeft = (viewport.scrollLeft + offsetX) * scaleRatio - offsetX;
        viewport.scrollTop = (viewport.scrollTop + offsetY) * scaleRatio - offsetY;
        pinchScale = nextScale;
        applyPinchScale();
      }, { passive: false });

      viewport.addEventListener('touchend', (event) => {
        if (drawModeEnabled) return;
        if (event.touches.length < 2) {
          pinchStartDistance = 0;
        }
        if (pinchScale <= 1.01) resetPinchZoom();
      }, { passive: true });

      async function renderPageSheet(pageNumber) {
        const page = await pdfDoc.getPage(pageNumber);
        const sheet = document.createElement('div');
        sheet.className = 'page-sheet';
        sheet.dataset.page = String(pageNumber);

        const frame = document.createElement('div');
        frame.className = 'page-frame';
        const canvas = document.createElement('canvas');
        const overlay = document.createElement('div');
        overlay.className = 'overlay';
        const drawLayer = document.createElement('div');
        drawLayer.className = 'draw-layer';
        frame.appendChild(canvas);
        frame.appendChild(overlay);
        frame.appendChild(drawLayer);
        sheet.appendChild(frame);
        wrap.appendChild(sheet);
        attachDrawHandlers(frame, pageNumber);
        frame.classList.toggle('draw-enabled', drawModeEnabled);

        const context = canvas.getContext('2d');
        const baseViewport = page.getViewport({ scale: 1 });
        const dpr = Math.min(window.devicePixelRatio || 1, highQuality ? 3 : 2);
        const maxCssWidth = window.innerWidth || 860;
        const fitScale = maxCssWidth / baseViewport.width;
        const viewportScale = Math.min(fitScale, highQuality ? 2.6 : 2.2);

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
        const cssWidth = viewportObj.width / dpr;
        const cssHeight = viewportObj.height / dpr;
        const cssScale = cssWidth / baseViewport.width;
        const meta = {
          cssWidth,
          cssHeight,
          cssScale,
          baseViewportHeight: baseViewport.height
        };
        pageMeta.set(pageNumber, meta);
        drawPageOverlay(pageNumber, overlay, meta);
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
          pageMeta.clear();
          for (let pageNumber = 1; pageNumber <= pdfDoc.numPages; pageNumber += 1) {
            await renderPageSheet(pageNumber);
          }
          resetPinchZoom();
          refreshAllOverlays();
          if (activeFieldId) {
            requestAnimationFrame(() => scrollToActiveField());
          }
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

      function handleCommand(raw) {
        try {
          const message = typeof raw === 'string' ? JSON.parse(raw) : raw;
          if (!message || !message.type) return;
          if (message.type === 'setDrawMode') {
            setDrawMode(Boolean(message.enabled));
          }
          if (message.type === 'setDraftRect') {
            setDraftRect(message);
          }
          if (message.type === 'setDraftRectEditable') {
            setDraftRectEditable(Boolean(message.enabled), message.bottomReserve);
          }
          if (message.type === 'scrollToDraftRect') {
            if (Number.isFinite(Number(message.bottomReserve))) {
              draftBottomReservePx = Math.max(0, Number(message.bottomReserve));
            }
            requestAnimationFrame(() => scrollToDraftRect());
          }
          if (message.type === 'clearDraftRect') {
            clearDraftRect();
          }
          if (message.type === 'setActive') {
            activeFieldId = String(message.fieldId || '');
            refreshAllOverlays();
            if (activeFieldId) {
              requestAnimationFrame(() => scrollToActiveField());
            }
          }
          if (message.type === 'resetZoom') {
            resetPinchZoom();
          }
          if (message.type === 'updateHighlights') {
            if (Array.isArray(message.highlights)) {
              highlights = message.highlights;
              refreshAllOverlays();
            }
          }
        } catch (error) {
          console.error('PDF preview command failed', error);
        }
      }

      window.__applyPreviewCommand = handleCommand;
      document.addEventListener('message', (event) => handleCommand(event.data));
      window.addEventListener('message', (event) => handleCommand(event.data));
      void boot();
    </script>
  </body>
</html>`;
}
