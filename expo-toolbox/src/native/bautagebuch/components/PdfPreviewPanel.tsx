import { useCallback, useEffect, useRef, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import * as FileSystem from 'expo-file-system/legacy';
import { WebView, type WebViewMessageEvent } from 'react-native-webview';

import { PrimaryButton } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';

type Props = {
  pdfPath: string | null;
  loading?: boolean;
  error?: string | null;
};

const PREVIEW_HTML = `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=4.0" />
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
    <style>
      * { box-sizing: border-box; }
      html, body { margin: 0; padding: 0; background: #f2f0eb; min-height: 100%; }
      #wrap { position: relative; min-height: 320px; display: flex; justify-content: center; padding: 8px 0; }
      canvas { display: block; max-width: 100%; height: auto; background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
    </style>
  </head>
  <body>
    <div id="wrap">
      <canvas id="canvas"></canvas>
    </div>
    <script>
      const pdfjsLib = window.pdfjsLib;
      if (!pdfjsLib) {
        throw new Error('pdf.js nicht geladen');
      }
      pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

      let pdfDoc = null;
      let currentPage = 1;
      let renderTask = null;

      function post(payload) {
        if (window.ReactNativeWebView && window.ReactNativeWebView.postMessage) {
          window.ReactNativeWebView.postMessage(JSON.stringify(payload));
        }
      }

      async function renderPage(pageNumber) {
        if (!pdfDoc) return;
        const safePage = Math.min(Math.max(pageNumber, 1), pdfDoc.numPages);
        currentPage = safePage;
        const page = await pdfDoc.getPage(safePage);
        const canvas = document.getElementById('canvas');
        const context = canvas.getContext('2d');
        const baseViewport = page.getViewport({ scale: 1 });
        const dpr = Math.min(window.devicePixelRatio || 1, 3);
        const maxCssWidth = Math.min(window.innerWidth || 860, 860);
        const fitScale = maxCssWidth / baseViewport.width;
        const scale = fitScale * dpr;
        const viewport = page.getViewport({ scale });
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        canvas.style.width = Math.ceil(viewport.width / dpr) + 'px';
        canvas.style.height = Math.ceil(viewport.height / dpr) + 'px';
        if (renderTask) {
          try { renderTask.cancel(); } catch (_) {}
        }
        renderTask = page.render({ canvasContext: context, viewport });
        await renderTask.promise;
        renderTask = null;
        post({ type: 'state', page: safePage, pageCount: pdfDoc.numPages, ready: true, error: null });
      }

      async function loadPdfBase64(base64, preferredPage) {
        try {
          const data = atob(base64);
          const bytes = new Uint8Array(data.length);
          for (let i = 0; i < data.length; i += 1) bytes[i] = data.charCodeAt(i);
          pdfDoc = await pdfjsLib.getDocument({ data: bytes }).promise;
          const targetPage = Math.min(Math.max(preferredPage || currentPage || 1, 1), pdfDoc.numPages);
          await renderPage(targetPage);
        } catch (error) {
          post({
            type: 'state',
            page: 1,
            pageCount: 1,
            ready: false,
            error: error && error.message ? error.message : 'PDF-Vorschau fehlgeschlagen'
          });
        }
      }

      function handleCommand(raw) {
        try {
          const message = typeof raw === 'string' ? JSON.parse(raw) : raw;
          if (!message || !message.type) return;
          if (message.type === 'setPage') {
            void renderPage(Number(message.page || 1));
            return;
          }
          if (message.type === 'setPdf' && message.base64) {
            void loadPdfBase64(String(message.base64), Number(message.page || currentPage || 1));
          }
        } catch {
          // Ignore malformed commands.
        }
      }

      document.addEventListener('message', (event) => handleCommand(event.data));
      window.addEventListener('message', (event) => handleCommand(event.data));
      post({ type: 'ready' });
    </script>
  </body>
</html>`;

export function PdfPreviewPanel({ pdfPath, loading = false, error = null }: Props) {
  const webViewRef = useRef<WebView>(null);
  const [webReady, setWebReady] = useState(false);
  const [readBusy, setReadBusy] = useState(false);
  const [renderError, setRenderError] = useState<string | null>(null);
  const [hasRendered, setHasRendered] = useState(false);
  const [previewState, setPreviewState] = useState({
    page: 1,
    pageCount: 1,
    ready: false
  });
  const pendingBase64Ref = useRef<string | null>(null);
  const currentPageRef = useRef(1);

  const pushPdfToWebView = useCallback((base64: string, page = currentPageRef.current) => {
    webViewRef.current?.postMessage(
      JSON.stringify({
        type: 'setPdf',
        base64,
        page
      })
    );
  }, []);

  useEffect(() => {
    let cancelled = false;
    if (!pdfPath) {
      pendingBase64Ref.current = null;
      setHasRendered(false);
      setPreviewState({ page: 1, pageCount: 1, ready: false });
      return undefined;
    }

    setReadBusy(true);
    setRenderError(null);
    void FileSystem.readAsStringAsync(pdfPath, {
      encoding: FileSystem.EncodingType.Base64
    })
      .then((base64) => {
        if (cancelled) return;
        pendingBase64Ref.current = base64;
        if (webReady) {
          pushPdfToWebView(base64, currentPageRef.current);
        }
      })
      .catch((readError) => {
        if (!cancelled) {
          pendingBase64Ref.current = null;
          setRenderError(readError instanceof Error ? readError.message : 'PDF konnte nicht gelesen werden.');
        }
      })
      .finally(() => {
        if (!cancelled) setReadBusy(false);
      });

    return () => {
      cancelled = true;
    };
  }, [pdfPath, pushPdfToWebView, webReady]);

  useEffect(() => {
    if (!webReady || !pendingBase64Ref.current) return;
    pushPdfToWebView(pendingBase64Ref.current, currentPageRef.current);
  }, [webReady, pushPdfToWebView]);

  const goToPage = (page: number) => {
    currentPageRef.current = page;
    webViewRef.current?.postMessage(JSON.stringify({ type: 'setPage', page }));
  };

  const onWebMessage = (event: WebViewMessageEvent) => {
    try {
      const payload = JSON.parse(event.nativeEvent.data) as {
        type?: string;
        page?: number;
        pageCount?: number;
        ready?: boolean;
        error?: string | null;
      };
      if (payload.type === 'ready') {
        setWebReady(true);
        return;
      }
      if (payload.type !== 'state') return;
      if (payload.error) {
        setRenderError(payload.error);
        return;
      }
      setRenderError(null);
      const page = Number(payload.page || 1);
      currentPageRef.current = page;
      setHasRendered(true);
      setPreviewState({
        page,
        pageCount: Number(payload.pageCount || 1),
        ready: Boolean(payload.ready)
      });
    } catch {
      setRenderError('PDF-Vorschau konnte nicht dargestellt werden.');
    }
  };

  const displayError = error || renderError;
  const showInitialLoader = !hasRendered && (loading || readBusy || !pdfPath);
  const showUpdateOverlay = hasRendered && (loading || readBusy);

  if (displayError && !hasRendered) {
    return (
      <View style={styles.center}>
        <Text style={styles.error}>{displayError}</Text>
      </View>
    );
  }

  if (showInitialLoader) {
    return (
      <View style={styles.center}>
        <ActivityIndicator color={colors.accent} />
        <Text style={styles.muted}>Live-PDF wird vorbereitet…</Text>
      </View>
    );
  }

  return (
    <View style={styles.root}>
      <Text style={styles.banner}>Live-PDF-Vorschau</Text>
      <View style={styles.panel}>
        <WebView
          ref={webViewRef}
          source={{ html: PREVIEW_HTML }}
          style={styles.webview}
          originWhitelist={['*']}
          scrollEnabled
          javaScriptEnabled
          domStorageEnabled
          onMessage={onWebMessage}
          onError={() => setRenderError('PDF-Vorschau konnte nicht geladen werden.')}
          onHttpError={() => setRenderError('PDF-Vorschau konnte nicht geladen werden.')}
        />
        {showUpdateOverlay ? (
          <View style={styles.updateOverlay} pointerEvents="none">
            <ActivityIndicator color={colors.accent} />
            <Text style={styles.updateLabel}>Aktualisiere…</Text>
          </View>
        ) : null}
        {displayError && hasRendered ? (
          <View style={styles.updateOverlay}>
            <Text style={styles.error}>{displayError}</Text>
          </View>
        ) : null}
      </View>
      <View style={styles.controls}>
        <PrimaryButton
          label="◀"
          variant="ghost"
          disabled={!previewState.ready || previewState.page <= 1}
          onPress={() => goToPage(previewState.page - 1)}
        />
        <Text style={styles.pageLabel}>
          Seite {previewState.page} / {previewState.pageCount}
        </Text>
        <PrimaryButton
          label="▶"
          variant="ghost"
          disabled={!previewState.ready || previewState.page >= previewState.pageCount}
          onPress={() => goToPage(previewState.page + 1)}
        />
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: { gap: spacing.sm },
  banner: { ...typography.label, color: colors.accent2 },
  panel: {
    minHeight: 360,
    borderRadius: 12,
    overflow: 'hidden',
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  webview: {
    flex: 1,
    minHeight: 360,
    backgroundColor: colors.panel
  },
  updateOverlay: {
    ...StyleSheet.absoluteFillObject,
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.xs,
    backgroundColor: 'rgba(242, 240, 235, 0.72)'
  },
  updateLabel: { ...typography.caption, color: colors.muted },
  controls: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm
  },
  pageLabel: { ...typography.caption, color: colors.muted, minWidth: 110, textAlign: 'center' },
  center: {
    minHeight: 200,
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm,
    padding: spacing.md,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  muted: { ...typography.caption, color: colors.muted, textAlign: 'center' },
  error: { ...typography.body, color: colors.danger, textAlign: 'center' }
});
