import { useEffect, useMemo, useRef, useState } from 'react';
import { ActivityIndicator, Dimensions, StyleSheet, Text, View } from 'react-native';
import * as FileSystem from 'expo-file-system/legacy';
import { WebView, type WebViewMessageEvent } from 'react-native-webview';

import { PrimaryButton } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';
import type { DetectedField } from '../types';
import { PdfPreviewPanel } from './PdfPreviewPanel';

type Props = {
  pdfPath: string | null;
  detectedFields?: DetectedField[];
  activeFieldId?: string | null;
  activeFieldLabel?: string | null;
  activeFieldPage?: number;
  variant?: 'default' | 'pinned';
};

const PINNED_PREVIEW_HEIGHT = Math.max(220, Math.round(Dimensions.get('window').height * 0.34));

type PreviewState = {
  page: number;
  pageCount: number;
  ready: boolean;
  error: string | null;
};

function buildPreviewHtml(base64: string, highlights: Array<{ fieldId: string; page: number; rect: number[] }>) {
  const highlightsJson = JSON.stringify(highlights);
  return `<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0" />
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
    <style>
      * { box-sizing: border-box; }
      html, body { margin: 0; padding: 0; background: #f2f0eb; height: 100%; }
      #wrap { position: relative; min-height: 320px; }
      canvas { display: block; width: 100%; height: auto; }
      #overlay { position: absolute; inset: 0; pointer-events: none; }
      .highlight {
        position: absolute;
        border: 2px solid rgba(47, 111, 237, 0.55);
        background: rgba(47, 111, 237, 0.12);
        border-radius: 4px;
      }
      .highlight.active {
        border-color: rgba(214, 69, 69, 0.9);
        background: rgba(214, 69, 69, 0.18);
        box-shadow: 0 0 0 2px rgba(214, 69, 69, 0.2);
      }
    </style>
  </head>
  <body>
    <div id="wrap">
      <canvas id="canvas"></canvas>
      <div id="overlay"></div>
    </div>
    <script>
      const pdfjsLib = window.pdfjsLib;
      if (!pdfjsLib) {
        throw new Error('pdf.js nicht geladen');
      }
      pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

      const highlights = ${highlightsJson};
      let pdfDoc = null;
      let currentPage = 1;
      let activeFieldId = '';
      let viewportScale = 1;

      function post(payload) {
        if (window.ReactNativeWebView && window.ReactNativeWebView.postMessage) {
          window.ReactNativeWebView.postMessage(JSON.stringify(payload));
        }
      }

      function pageHighlights(pageNumber) {
        return highlights.filter((entry) => Number(entry.page || 1) === pageNumber);
      }

      function drawOverlay(pageNumber, viewport) {
        const overlay = document.getElementById('overlay');
        overlay.innerHTML = '';
        overlay.style.width = viewport.width + 'px';
        overlay.style.height = viewport.height + 'px';

        for (const entry of pageHighlights(pageNumber)) {
          if (!Array.isArray(entry.rect) || entry.rect.length < 4) continue;
          const [x1, y1, x2, y2] = entry.rect;
          const left = Math.min(x1, x2) * viewportScale;
          const top = (viewport.height / viewportScale - Math.max(y1, y2)) * viewportScale;
          const width = Math.abs(x2 - x1) * viewportScale;
          const height = Math.abs(y2 - y1) * viewportScale;
          const box = document.createElement('div');
          box.className = 'highlight' + (entry.fieldId === activeFieldId ? ' active' : '');
          box.style.left = left + 'px';
          box.style.top = top + 'px';
          box.style.width = width + 'px';
          box.style.height = height + 'px';
          overlay.appendChild(box);
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
        const maxWidth = Math.min(window.innerWidth || 860, 860);
        viewportScale = Math.min(maxWidth / baseViewport.width, 2.1);
        const viewport = page.getViewport({ scale: viewportScale });
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        await page.render({ canvasContext: context, viewport }).promise;
        drawOverlay(safePage, viewport);
        post({ type: 'state', page: safePage, pageCount: pdfDoc.numPages, ready: true, error: null });
      }

      async function boot() {
        try {
          const data = atob('${base64}');
          const bytes = new Uint8Array(data.length);
          for (let i = 0; i < data.length; i += 1) bytes[i] = data.charCodeAt(i);
          pdfDoc = await pdfjsLib.getDocument({ data: bytes }).promise;
          await renderPage(1);
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
          }
          if (message.type === 'setActive') {
            activeFieldId = String(message.fieldId || '');
            const page = Number(message.page || currentPage || 1);
            void renderPage(page);
          }
        } catch {
          // Ignore malformed commands.
        }
      }

      document.addEventListener('message', (event) => handleCommand(event.data));
      window.addEventListener('message', (event) => handleCommand(event.data));
      void boot();
    </script>
  </body>
</html>`;
}

export function SetupPdfFieldPreview({
  pdfPath,
  detectedFields = [],
  activeFieldId,
  activeFieldLabel,
  activeFieldPage = 1,
  variant = 'default'
}: Props) {
  const pinned = variant === 'pinned';
  const webViewRef = useRef<WebView>(null);
  const [html, setHtml] = useState<string | null>(null);
  const [readBusy, setReadBusy] = useState(false);
  const [useFallback, setUseFallback] = useState(false);
  const [previewState, setPreviewState] = useState<PreviewState>({
    page: 1,
    pageCount: 1,
    ready: false,
    error: null
  });

  const highlights = useMemo(
    () =>
      detectedFields
        .filter((field) => Array.isArray(field.rect) && field.rect.length >= 4)
        .map((field) => ({
          fieldId: field.fieldId,
          page: Number(field.page || 1),
          rect: field.rect as number[]
        })),
    [detectedFields]
  );

  useEffect(() => {
    let cancelled = false;
    if (!pdfPath) {
      setHtml(null);
      setUseFallback(false);
      return undefined;
    }

    setReadBusy(true);
    setUseFallback(false);
    void FileSystem.readAsStringAsync(pdfPath, {
      encoding: FileSystem.EncodingType.Base64
    })
      .then((base64) => {
        if (cancelled) return;
        setHtml(buildPreviewHtml(base64, highlights));
        setPreviewState({ page: 1, pageCount: 1, ready: false, error: null });
      })
      .catch(() => {
        if (!cancelled) {
          setHtml(null);
          setUseFallback(true);
        }
      })
      .finally(() => {
        if (!cancelled) setReadBusy(false);
      });

    return () => {
      cancelled = true;
    };
  }, [pdfPath, highlights]);

  useEffect(() => {
    if (!previewState.ready || useFallback) return;
    webViewRef.current?.postMessage(
      JSON.stringify({
        type: 'setActive',
        fieldId: activeFieldId || '',
        page: Number(activeFieldPage || previewState.page || 1)
      })
    );
  }, [activeFieldId, activeFieldPage, previewState.ready, useFallback]);

  const goToPage = (page: number) => {
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
      if (payload.type !== 'state') return;
      if (payload.error) {
        setUseFallback(true);
        return;
      }
      setPreviewState({
        page: Number(payload.page || 1),
        pageCount: Number(payload.pageCount || 1),
        ready: Boolean(payload.ready),
        error: payload.error || null
      });
    } catch {
      setUseFallback(true);
    }
  };

  if (readBusy) {
    return (
      <View style={styles.center}>
        <ActivityIndicator color={colors.accent} />
        <Text style={styles.muted}>Vorlagen-PDF wird geladen…</Text>
      </View>
    );
  }

  if (useFallback || !html) {
    return <PdfPreviewPanel pdfPath={pdfPath} error={previewState.error} />;
  }

  const banner = activeFieldLabel
    ? `Aktiv: ${activeFieldLabel}${activeFieldPage ? ` · Seite ${activeFieldPage}` : ''}`
    : pinned
      ? 'Feld oder Spalte antippen — Markierung in der PDF zeigt die Position.'
      : 'Feld in der Liste antippen, um die zugehörige PDF-Seite zu sehen.';

  return (
    <View style={[styles.root, pinned ? styles.rootPinned : null]}>
      <Text style={[styles.banner, pinned ? styles.bannerPinned : null]} numberOfLines={2}>
        {banner}
      </Text>
      <View style={[styles.panel, pinned ? styles.panelPinned : null]}>
        <WebView
          ref={webViewRef}
          source={{ html }}
          style={[styles.webview, pinned ? styles.webviewPinned : null]}
          originWhitelist={['*']}
          scrollEnabled
          onMessage={onWebMessage}
          onError={() => setUseFallback(true)}
          onHttpError={() => setUseFallback(true)}
        />
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
      {highlights.length === 0 && !pinned ? (
        <Text style={styles.hint}>
          Feld-Overlays sind ohne Positionsdaten nicht verfügbar. Seitennavigation und Feld-Banner funktionieren weiterhin.
        </Text>
      ) : null}
      {pinned ? (
        <Text style={styles.legend}>Markierung: aktives Feld · Seite mit Pfeilen wechseln</Text>
      ) : null}
    </View>
  );
}

const styles = StyleSheet.create({
  root: { gap: spacing.sm },
  rootPinned: {
    gap: spacing.xs,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    paddingBottom: spacing.xs
  },
  banner: { ...typography.caption, color: colors.accent2 },
  bannerPinned: {
    ...typography.label,
    color: colors.ink
  },
  panel: {
    minHeight: 360,
    borderRadius: 12,
    overflow: 'hidden',
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  panelPinned: {
    minHeight: PINNED_PREVIEW_HEIGHT,
    height: PINNED_PREVIEW_HEIGHT,
    borderRadius: 10
  },
  webview: {
    flex: 1,
    minHeight: 360,
    backgroundColor: colors.panel
  },
  webviewPinned: {
    minHeight: PINNED_PREVIEW_HEIGHT - 2,
    height: PINNED_PREVIEW_HEIGHT - 2
  },
  controls: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm
  },
  pageLabel: { ...typography.caption, color: colors.muted, minWidth: 110, textAlign: 'center' },
  hint: { ...typography.caption, color: colors.muted },
  legend: {
    ...typography.caption,
    color: colors.muted,
    paddingBottom: spacing.xxs
  },
  center: {
    minHeight: 200,
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm,
    padding: spacing.md
  },
  muted: { ...typography.caption, color: colors.muted, textAlign: 'center' }
});
