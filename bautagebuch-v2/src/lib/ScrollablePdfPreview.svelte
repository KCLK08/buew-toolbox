<script>
  import { onDestroy } from 'svelte';
  import { loadPdfPreviewDocument } from '$lib/pdf-scan.js';

  /** @type {Uint8Array | null} */
  export let pdfBytes = null;
  export let reloadKey = 0;
  export let loading = false;
  export let error = '';

  const PDF_PREVIEW_LOAD_ERROR = 'PDF Vorschau konnte nicht geladen werden';
  const MAX_CANVAS_PIXELS = 12_000_000;

  const nodes = { viewport: null, wrap: null };
  const gesture = { scale: 1, startDistance: 0, startScale: 1 };
  let rendering = false;
  let renderError = '';
  let renderToken = 0;
  let renderedStamp = '';

  function applyPinchScale() {
    const wrap = nodes.wrap;
    if (!wrap) return;
    wrap.style.transform = `scale(${gesture.scale})`;
    wrap.style.transformOrigin = 'center top';
  }

  function resetPinchZoom() {
    gesture.scale = 1;
    gesture.startDistance = 0;
    gesture.startScale = 1;
    applyPinchScale();
  }

  function touchDistance(touches) {
    const dx = touches[0].clientX - touches[1].clientX;
    const dy = touches[0].clientY - touches[1].clientY;
    return Math.hypot(dx, dy);
  }

  function bindPinch(node) {
    const onTouchStart = (event) => {
      if (event.touches.length === 2) {
        gesture.startDistance = touchDistance(event.touches);
        gesture.startScale = gesture.scale;
      }
    };
    const onTouchMove = (event) => {
      if (event.touches.length !== 2 || gesture.startDistance <= 0) return;
      event.preventDefault();
      const ratio = touchDistance(event.touches) / gesture.startDistance;
      gesture.scale = Math.min(4, Math.max(1, gesture.startScale * ratio));
      applyPinchScale();
    };
    const onTouchEnd = () => {
      if (gesture.scale <= 1.01) resetPinchZoom();
    };
    node.addEventListener('touchstart', onTouchStart, { passive: true });
    node.addEventListener('touchmove', onTouchMove, { passive: false });
    node.addEventListener('touchend', onTouchEnd, { passive: true });
    return {
      destroy() {
        node.removeEventListener('touchstart', onTouchStart);
        node.removeEventListener('touchmove', onTouchMove);
        node.removeEventListener('touchend', onTouchEnd);
      }
    };
  }

  function mountViewport(node) {
    nodes.viewport = node;
    requestRender(pdfBytes, reloadKey);
    return {
      destroy() {
        if (nodes.viewport === node) nodes.viewport = null;
      }
    };
  }

  function mountWrap(node) {
    nodes.wrap = node;
    requestRender(pdfBytes, reloadKey);
    return {
      destroy() {
        if (nodes.wrap === node) nodes.wrap = null;
      }
    };
  }

  function previewStamp(bytes, revision) {
    if (!bytes || bytes.length === 0) return `empty:${revision}`;
    const last = bytes.length - 1;
    return `${revision}:${bytes.length}:${bytes[0] || 0}:${bytes[last >> 1] || 0}:${bytes[last] || 0}`;
  }

  function requestRender(bytes, revision) {
    if (!nodes.viewport || !nodes.wrap) return;
    void renderPages(bytes, revision);
  }

  async function renderPages(bytes, revision) {
    const stamp = previewStamp(bytes, revision);
    if (stamp === renderedStamp && nodes.wrap?.childElementCount) return;

    const token = ++renderToken;
    const wrap = nodes.wrap;
    const viewport = nodes.viewport;
    renderError = '';
    if (!viewport || !wrap) return;

    if (!bytes || bytes.length === 0) {
      wrap.replaceChildren();
      renderedStamp = stamp;
      rendering = false;
      resetPinchZoom();
      return;
    }

    rendering = true;
    try {
      const copy = bytes instanceof Uint8Array ? bytes.slice() : new Uint8Array(bytes);
      const blob = new Blob([copy], { type: 'application/pdf' });
      const doc = await loadPdfPreviewDocument(blob);
      if (token !== renderToken || nodes.wrap !== wrap) return;

      wrap.replaceChildren();
      const dpr = Math.min(window.devicePixelRatio || 1, 3);
      const maxCssWidth = Math.max(240, (viewport.clientWidth || window.innerWidth || 860) - 16);

      for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber += 1) {
        if (token !== renderToken || nodes.wrap !== wrap) return;
        const page = await doc.getPage(pageNumber);
        const baseViewport = page.getViewport({ scale: 1 });
        const fitScale = maxCssWidth / baseViewport.width;
        let pixelScale = Math.min(fitScale, 2.6) * dpr;
        const pixelArea = baseViewport.width * baseViewport.height * pixelScale * pixelScale;
        if (pixelArea > MAX_CANVAS_PIXELS) {
          pixelScale = Math.sqrt(MAX_CANVAS_PIXELS / (baseViewport.width * baseViewport.height));
        }

        const pageViewport = page.getViewport({ scale: pixelScale });
        const canvas = document.createElement('canvas');
        canvas.width = Math.ceil(pageViewport.width);
        canvas.height = Math.ceil(pageViewport.height);
        canvas.style.width = `${Math.ceil(pageViewport.width / dpr)}px`;
        canvas.style.height = `${Math.ceil(pageViewport.height / dpr)}px`;

        const context = canvas.getContext('2d');
        if (!context) continue;

        const sheet = document.createElement('div');
        sheet.className = 'page-sheet';
        sheet.appendChild(canvas);
        wrap.appendChild(sheet);

        await page.render({ canvasContext: context, viewport: pageViewport }).promise;
      }

      if (token !== renderToken) return;
      renderedStamp = stamp;
      resetPinchZoom();
    } catch (renderFailure) {
      if (token !== renderToken) return;
      console.error('PDF preview render failed', renderFailure);
      renderError = PDF_PREVIEW_LOAD_ERROR;
      wrap.replaceChildren();
      renderedStamp = '';
    } finally {
      if (token === renderToken) rendering = false;
    }
  }

  $: requestRender(pdfBytes, reloadKey);

  onDestroy(() => {
    renderToken += 1;
  });

  $: displayError = error || renderError;
</script>

<div class="scrollable-pdf-preview">
  {#if displayError && (!pdfBytes || pdfBytes.length === 0)}
    <p class="preview-state preview-state--error">{displayError}</p>
  {:else if !pdfBytes || pdfBytes.length === 0}
    <div class="preview-state">
      <p class="meta">{loading ? 'Live-PDF wird vorbereitet…' : 'Noch keine Live-Vorschau verfügbar.'}</p>
    </div>
  {:else}
    <div class="preview-viewport" use:mountViewport use:bindPinch>
      <div class="preview-wrap" use:mountWrap></div>
    </div>
    <p class="zoom-hint">Scrollen · Zwei Finger zum Zoomen</p>
    {#if loading || rendering}
      <div class="preview-overlay" aria-hidden="true">
        <p class="meta">{loading ? 'Aktualisiere…' : 'Rendert…'}</p>
      </div>
    {/if}
    {#if displayError}
      <div class="preview-overlay preview-overlay--error">
        <p>{displayError}</p>
      </div>
    {/if}
  {/if}
</div>

<style>
  .scrollable-pdf-preview {
    background: #f2f0eb;
    display: flex;
    flex: 1;
    flex-direction: column;
    height: 100%;
    min-height: 0;
    position: relative;
    width: 100%;
  }

  .preview-state {
    align-items: center;
    background: #f2f0eb;
    display: flex;
    flex: 1;
    justify-content: center;
    min-height: 200px;
    padding: 16px;
    text-align: center;
  }

  .preview-state--error {
    color: #8b3a32;
  }

  .preview-viewport {
    flex: 1;
    min-height: 0;
    overflow: auto;
    touch-action: pan-x pan-y;
    -webkit-overflow-scrolling: touch;
    width: 100%;
  }

  .preview-wrap {
    align-items: center;
    display: flex;
    flex-direction: column;
    gap: 12px;
    padding: 8px 0 16px;
    transform-origin: center top;
    width: 100%;
  }

  :global(.scrollable-pdf-preview .page-sheet) {
    display: flex;
    justify-content: center;
    width: 100%;
  }

  :global(.scrollable-pdf-preview canvas) {
    background: #fff;
    box-shadow: 0 1px 6px rgba(0, 0, 0, 0.06);
    display: block;
    height: auto;
    max-width: 100%;
  }

  .zoom-hint {
    background: rgba(255, 255, 255, 0.88);
    border-radius: 8px;
    bottom: 8px;
    color: #666;
    font-size: 11px;
    line-height: 1.3;
    margin: 0;
    opacity: 0.85;
    padding: 4px 8px;
    pointer-events: none;
    position: absolute;
    right: 8px;
  }

  .preview-overlay {
    align-items: center;
    background: rgba(242, 240, 235, 0.72);
    display: flex;
    inset: 0;
    justify-content: center;
    pointer-events: none;
    position: absolute;
  }

  .preview-overlay--error {
    background: rgba(242, 240, 235, 0.92);
    color: #8b3a32;
    padding: 12px;
    pointer-events: auto;
    text-align: center;
  }
</style>
