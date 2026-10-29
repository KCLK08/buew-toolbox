<script>
  import { onDestroy, onMount } from 'svelte';
  import { buildSimplePdfPreviewHtml, PDF_PREVIEW_LOAD_ERROR } from '$lib/pdf-preview-html.js';
  import { loadPdfPreviewAssets } from '$lib/pdf-preview-assets.js';

  /** @type {Uint8Array | null} */
  export let pdfBytes = null;
  export let reloadKey = 0;
  export let loading = false;
  export let error = '';

  let html = null;
  let readBusy = false;
  let renderError = '';
  let ready = false;
  /** @type {HTMLIFrameElement | null} */
  let iframeEl = null;

  function uint8ToBase64(bytes) {
    const view = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    let binary = '';
    const chunkSize = 0x8000;
    for (let offset = 0; offset < view.length; offset += chunkSize) {
      binary += String.fromCharCode(...view.subarray(offset, offset + chunkSize));
    }
    return btoa(binary);
  }

  async function rebuildHtml(bytes, revision) {
    void revision;
    if (!bytes || bytes.length === 0) {
      html = null;
      renderError = '';
      ready = false;
      return;
    }

    readBusy = true;
    renderError = '';
    ready = false;
    try {
      const assets = await loadPdfPreviewAssets();
      const base64 = uint8ToBase64(bytes);
      html = buildSimplePdfPreviewHtml({ base64, ...assets });
    } catch (readError) {
      console.error('PDF preview setup failed', readError);
      html = null;
      renderError = PDF_PREVIEW_LOAD_ERROR;
    } finally {
      readBusy = false;
    }
  }

  $: void rebuildHtml(pdfBytes, reloadKey);

  function onWindowMessage(event) {
    if (!iframeEl?.contentWindow || event.source !== iframeEl.contentWindow) return;
    try {
      const payload = JSON.parse(String(event.data || ''));
      if (payload?.type !== 'state') return;
      if (payload.error) {
        renderError = PDF_PREVIEW_LOAD_ERROR;
        ready = false;
        return;
      }
      renderError = '';
      ready = Boolean(payload.ready);
    } catch (messageError) {
      console.error('PDF preview message parse failed', messageError);
      renderError = PDF_PREVIEW_LOAD_ERROR;
      ready = false;
    }
  }

  onMount(() => {
    window.addEventListener('message', onWindowMessage);
  });

  onDestroy(() => {
    window.removeEventListener('message', onWindowMessage);
  });

  $: displayError = error || renderError;
  $: showInitialLoader = !html && (loading || readBusy || !pdfBytes?.length);
  $: showUpdateOverlay = Boolean(html) && (loading || readBusy);
</script>

<div class="scrollable-pdf-preview">
  {#if displayError && !html}
    <p class="preview-state preview-state--error">{displayError}</p>
  {:else if showInitialLoader}
    <div class="preview-state">
      <p class="meta">Live-PDF wird vorbereitet…</p>
    </div>
  {:else if !html}
    <p class="preview-state meta">Noch keine Live-Vorschau verfügbar.</p>
  {:else}
    <div class="preview-frame-wrap">
      <iframe bind:this={iframeEl} title="Live-PDF-Vorschau" srcdoc={html} class="preview-frame"></iframe>
      {#if showUpdateOverlay}
        <div class="preview-overlay" aria-hidden="true">
          <p class="meta">Aktualisiere…</p>
        </div>
      {/if}
      {#if displayError}
        <div class="preview-overlay preview-overlay--error">
          <p>{displayError}</p>
        </div>
      {/if}
      {#if !ready && !displayError && !showUpdateOverlay}
        <div class="preview-overlay" aria-hidden="true">
          <p class="meta">Rendert…</p>
        </div>
      {/if}
    </div>
  {/if}
</div>

<style>
  .scrollable-pdf-preview {
    display: flex;
    flex: 1;
    flex-direction: column;
    min-height: 0;
    width: 100%;
  }

  .preview-state {
    align-items: center;
    background: #fffefb;
    border: 1px solid var(--border, #ddd4c8);
    border-radius: 12px;
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

  .preview-frame-wrap {
    border: 1px solid var(--border, #ddd4c8);
    border-radius: 12px;
    flex: 1;
    min-height: 0;
    overflow: hidden;
    position: relative;
  }

  .preview-frame {
    background: #f2f0eb;
    border: 0;
    display: block;
    height: 100%;
    min-height: 320px;
    width: 100%;
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
