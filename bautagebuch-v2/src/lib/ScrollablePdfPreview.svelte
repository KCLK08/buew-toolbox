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

  /** @type {HTMLDivElement | null} */
  let viewportEl = null;
  /** @type {HTMLDivElement | null} */
  let wrapEl = null;
  let rendering = false;
  let renderError = '';
  let pinchScale = 1;
  let pinchStartDistance = 0;
  let pinchStartScale = 1;
  let renderToken = 0;

  function resetPinchZoom() {
    pinchScale = 1;
    pinchStartDistance = 0;
    if (wrapEl) wrapEl.style.transform = 'scale(1)';
  }

  function touchDistance(touches) {
    const dx = touches[0].clientX - touches[1].clientX;
    const dy = touches[0].clientY - touches[1].clientY;
    return Math.hypot(dx, dy);
  }

  function applyPinchScale() {
    if (!wrapEl) return;
    wrapEl.style.transform = `scale(${pinchScale})`;
    wrapEl.style.transformOrigin = 'center top';
  }

  function bindPinch(node) {
    const onTouchStart = (event) => {
      if (event.touches.length === 2) {
        pinchStartDistance = touchDistance(event.touches);
        pinchStartScale = pinchScale;
      }
    };
    const onTouchMove = (event) => {
      if (event.touches.length !== 2 || pinchStartDistance <= 0) return;
      event.preventDefault();
      const ratio = touchDistance(event.touches) / pinchStartDistance;
      pinchScale = Math.min(4, Math.max(1, pinchStartScale * ratio));
      applyPinchScale();
    };
    const onTouchEnd = () => {
      if (pinchScale <= 1.01) resetPinchZoom();
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

  async function renderPages(bytes, revision) {
    void revision;
    const token = ++renderToken;
    renderError = '';
    resetPinchZoom();

    if (!viewportEl || !wrapEl) return;
    if (!bytes || bytes.length === 0) {
      wrapEl.replaceChildren();
      rendering = false;
      return;
    }

    rendering = true;
    try {
      const copy = bytes instanceof Uint8Array ? bytes.slice() : new Uint8Array(bytes);
      const blob = new Blob([copy], { type: 'application/pdf' });
      const doc = await loadPdfPreviewDocument(blob);
      if (token !== renderToken || !wrapEl) return;

      wrapEl.replaceChildren();
      const dpr = Math.min(window.devicePixelRatio || 1, 3);
      const maxCssWidth = Math.max(240, (viewportEl?.clientWidth || window.innerWidth || 860) - 16);

      for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber += 1) {
        if (token !== renderToken || !wrapEl) return;
        const page = await doc.getPage(pageNumber);
        const baseViewport = page.getViewport({ scale: 1 });
        const fitScale = maxCssWidth / baseViewport.width;
        const viewportScale = Math.min(fitScale, 2.6);
        let pixelScale = viewportScale * dpr;
        const pixelArea = baseViewport.width * baseViewport.height * pixelScale * pixelScale;
        if (pixelArea > MAX_CANVAS_PIXELS) {
          pixelScale = Math.sqrt(MAX_CANVAS_PIXELS / (baseViewport.width * baseViewport.height));
        }

        const viewport = page.getViewport({ scale: pixelScale });
        const canvas = document.createElement('canvas');
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        canvas.style.width = `${Math.ceil(viewport.width / dpr)}px`;
        canvas.style.height = `${Math.ceil(viewport.height / dpr)}px`;

        const context = canvas.getContext('2d');
        if (!context) continue;

        const sheet = document.createElement('div');
        sheet.className = 'page-sheet';
        sheet.appendChild(canvas);
        wrapEl.appendChild(sheet);

        await page.render({ canvasContext: context, viewport }).promise;
      }
    } catch (renderFailure) {
      if (token !== renderToken) return;
      console.error('PDF preview render failed', renderFailure);
      renderError = PDF_PREVIEW_LOAD_ERROR;
      wrapEl?.replaceChildren();
    } finally {
      if (token === renderToken) rendering = false;
    }
  }

  $: if (viewportEl && wrapEl) {
    void renderPages(pdfBytes, reloadKey);
  }

  onDestroy(() => {
    renderToken += 1;
  });

  $: displayError = error || renderError;
  $: showInitialLoader = !displayError && (loading || rendering) && (!pdfBytes || pdfBytes.length === 0 || rendering);
</script>

<div class="scrollable-pdf-preview">
  {#if displayError && (!pdfBytes || pdfBytes.length === 0)}
    <p class="preview-state preview-state--error">{displayError}</p>
  {:else if !pdfBytes || pdfBytes.length === 0}
    <div class="preview-state">
      <p class="meta">{loading ? 'Live-PDF wird vorbereitet…' : 'Noch keine Live-Vorschau verfügbar.'}</p>
    </div>
  {:else}
    <div class="preview-viewport" bind:this={viewportEl} use:bindPinch>
      <div class="preview-wrap" bind:this={wrapEl}></div>
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
