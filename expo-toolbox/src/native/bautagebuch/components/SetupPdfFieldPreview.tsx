import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ActivityIndicator, Dimensions, StyleSheet, Text, View } from 'react-native';
import * as FileSystem from 'expo-file-system/legacy';
import { WebView, type WebViewMessageEvent } from 'react-native-webview';

import { PrimaryButton } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';
import { loadPdfPreviewAssets } from '../lib/pdf-preview-assets';
import {
  buildFieldPreviewHtml,
  buildScrollableFieldPreviewHtml,
  PDF_PREVIEW_LOAD_ERROR,
  type PreviewHtmlMode
} from '../lib/pdf-preview-html';
import {
  previewScrollOverlayPlacement,
  resolvePreviewOverlayPlacement
} from '../lib/pdf-preview-overlay';
import type { DetectedField } from '../types';
import { PdfPreviewPanel } from './PdfPreviewPanel';

type Props = {
  pdfPath: string | null;
  detectedFields?: DetectedField[];
  activeFieldId?: string | null;
  activeFieldLabel?: string | null;
  activeFieldPage?: number;
  variant?: 'default' | 'pinned' | 'mapping' | 'overlay';
  emphasizeActiveHighlight?: boolean;
};

const PINNED_PREVIEW_HEIGHT = Math.max(220, Math.round(Dimensions.get('window').height * 0.34));

type PreviewState = {
  page: number;
  pageCount: number;
  ready: boolean;
  error: string | null;
  renderMs?: number;
};

export function SetupPdfFieldPreview({
  pdfPath,
  detectedFields = [],
  activeFieldId,
  activeFieldLabel,
  activeFieldPage = 1,
  variant = 'default',
  emphasizeActiveHighlight = false
}: Props) {
  const pinned = variant === 'pinned';
  const mapping = variant === 'mapping';
  const overlay = variant === 'overlay';
  const highQuality = overlay || mapping;
  const highlightActive = emphasizeActiveHighlight || mapping || overlay;
  const webViewRef = useRef<WebView>(null);
  const [html, setHtml] = useState<string | null>(null);
  const [readBusy, setReadBusy] = useState(false);
  const [useFallback, setUseFallback] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);
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
          fieldId: String(field.fieldId),
          fieldName: String(field.fieldName || ''),
          page: Number(field.page || 1),
          rect: field.rect as number[]
        })),
    [detectedFields]
  );

  const resolvedActiveFieldId = useMemo(() => {
    if (!activeFieldId) return '';
    const needle = String(activeFieldId);
    const direct = detectedFields.find((field) => String(field.fieldId) === needle);
    if (direct) return String(direct.fieldId);
    const byName = detectedFields.find((field) => String(field.fieldName || '') === needle);
    if (byName) return String(byName.fieldId);
    return needle;
  }, [activeFieldId, detectedFields]);

  const activeFieldRect = useMemo(() => {
    if (!activeFieldId) return null;
    const match = detectedFields.find((field) => field.fieldId === activeFieldId);
    return Array.isArray(match?.rect) ? match.rect : null;
  }, [activeFieldId, detectedFields]);

  const overlayPlacement = useMemo(
    () => previewScrollOverlayPlacement(resolvePreviewOverlayPlacement(activeFieldRect)),
    [activeFieldRect]
  );

  const htmlMode: PreviewHtmlMode = mapping ? 'mapping' : overlay ? 'overlay' : pinned ? 'pinned' : 'default';

  useEffect(() => {
    let cancelled = false;
    if (!pdfPath) {
      setHtml(null);
      setUseFallback(false);
      return undefined;
    }

    setReadBusy(true);
    setUseFallback(false);
    setLoadError(null);
    void Promise.all([
      FileSystem.readAsStringAsync(pdfPath, {
        encoding: FileSystem.EncodingType.Base64
      }),
      loadPdfPreviewAssets()
    ])
      .then(([base64, assets]) => {
        if (cancelled) return;
        const previewOptions = {
          base64,
          ...assets,
          highlights,
          mode: htmlMode,
          highlightActive,
          highQuality
        };
        setHtml(
          overlay
            ? buildScrollableFieldPreviewHtml(previewOptions)
            : buildFieldPreviewHtml(previewOptions)
        );
        setPreviewState({ page: 1, pageCount: 1, ready: false, error: null });
      })
      .catch((error) => {
        console.error('PDF preview setup failed', error);
        if (!cancelled) {
          setHtml(null);
          setLoadError(PDF_PREVIEW_LOAD_ERROR);
          setUseFallback(true);
        }
      })
      .finally(() => {
        if (!cancelled) setReadBusy(false);
      });

    return () => {
      cancelled = true;
    };
  }, [pdfPath, highlights, htmlMode, highlightActive, highQuality, overlay]);

  const postPreviewCommand = useCallback((payload: Record<string, unknown>) => {
    const json = JSON.stringify(payload);
    webViewRef.current?.postMessage(json);
    webViewRef.current?.injectJavaScript(
      `(function(){try{if(window.__applyPreviewCommand){window.__applyPreviewCommand(${JSON.stringify(json)});}}catch(e){}})();true;`
    );
  }, []);

  useEffect(() => {
    if (!previewState.ready || useFallback) return;
    postPreviewCommand({
      type: 'setActive',
      fieldId: resolvedActiveFieldId,
      page: Number(activeFieldPage || previewState.page || 1),
      overlayPlacement
    });
  }, [
    resolvedActiveFieldId,
    activeFieldPage,
    overlayPlacement,
    previewState.ready,
    useFallback,
    postPreviewCommand
  ]);

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
        renderMs?: number;
      };
      if (payload.type !== 'state') return;
      if (payload.error) {
        console.error('PDF preview render error', payload.error);
        setLoadError(PDF_PREVIEW_LOAD_ERROR);
        setUseFallback(true);
        return;
      }
      setPreviewState({
        page: Number(payload.page || 1),
        pageCount: Number(payload.pageCount || 1),
        ready: Boolean(payload.ready),
        error: payload.error || null,
        renderMs: payload.renderMs
      });
    } catch (error) {
      console.error('PDF preview message parse failed', error);
      setLoadError(PDF_PREVIEW_LOAD_ERROR);
      setUseFallback(true);
    }
  };

  if (readBusy) {
    return (
      <View style={[styles.center, overlay ? styles.centerOverlay : null]}>
        <ActivityIndicator color={colors.accent} />
        <Text style={styles.muted}>Vorlagen-PDF wird geladen…</Text>
      </View>
    );
  }

  if (useFallback || !html) {
    return <PdfPreviewPanel pdfPath={pdfPath} error={loadError || previewState.error} />;
  }

  const showPageControls = !mapping && !overlay;
  const showBanner = !mapping;

  const banner = activeFieldLabel
    ? `Aktiv: ${activeFieldLabel}${activeFieldPage ? ` · Seite ${activeFieldPage}` : ''}`
    : overlay
      ? 'Feld in der Liste auswählen — die Markierung zeigt die PDF-Position.'
      : pinned
        ? 'Feld oder Spalte antippen — Markierung in der PDF zeigt die Position.'
        : 'Feld in der Liste antippen, um die zugehörige PDF-Seite zu sehen.';

  return (
    <View
      style={[
        styles.root,
        pinned ? styles.rootPinned : null,
        mapping ? styles.rootMapping : null,
        overlay ? styles.rootOverlay : null
      ]}
    >
      {showBanner ? (
        <Text
          style={[
            styles.banner,
            pinned || overlay ? styles.bannerPinned : null,
            overlay ? styles.bannerOverlay : null
          ]}
          numberOfLines={2}
        >
          {banner}
        </Text>
      ) : null}
      <View
        style={[
          styles.panel,
          pinned ? styles.panelPinned : null,
          mapping ? styles.panelMapping : null,
          overlay ? styles.panelOverlay : null
        ]}
      >
        <WebView
          ref={webViewRef}
          source={{ html }}
          style={[
            styles.webview,
            pinned ? styles.webviewPinned : null,
            mapping ? styles.webviewMapping : null,
            overlay ? styles.webviewOverlay : null
          ]}
          originWhitelist={['*']}
          scrollEnabled={false}
          bounces={false}
          overScrollMode="never"
          javaScriptEnabled
          domStorageEnabled
          allowsInlineMediaPlayback
          setBuiltInZoomControls={false}
          onMessage={onWebMessage}
          onLoadEnd={() => {
            postPreviewCommand({
              type: 'setActive',
              fieldId: resolvedActiveFieldId,
              page: Number(activeFieldPage || previewState.page || 1),
              overlayPlacement
            });
          }}
          onError={(event) => {
            console.error('PDF preview WebView error', event.nativeEvent);
            setLoadError(PDF_PREVIEW_LOAD_ERROR);
            setUseFallback(true);
          }}
          onHttpError={(event) => {
            console.error('PDF preview WebView HTTP error', event.nativeEvent);
            setLoadError(PDF_PREVIEW_LOAD_ERROR);
            setUseFallback(true);
          }}
        />
      </View>
      {showPageControls ? (
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
      ) : null}
      {highlights.length === 0 && !pinned && !mapping && !overlay ? (
        <Text style={styles.hint}>
          Feld-Overlays sind ohne Positionsdaten nicht verfügbar. Seitennavigation und Feld-Banner funktionieren
          weiterhin.
        </Text>
      ) : null}
      {pinned && !mapping ? (
        <Text style={styles.legend}>Markierung: aktives Feld · Seite mit Pfeilen wechseln · Zwei Finger zum Zoomen</Text>
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
  rootMapping: {
    flex: 1,
    gap: 0,
    paddingHorizontal: 0,
    paddingTop: 0,
    minHeight: 0
  },
  rootOverlay: {
    flex: 1,
    minHeight: 0,
    gap: spacing.xxs
  },
  banner: { ...typography.caption, color: colors.accent2 },
  bannerPinned: {
    ...typography.label,
    color: colors.ink
  },
  bannerOverlay: {
    paddingHorizontal: spacing.sm,
    paddingTop: spacing.xxs
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
  panelMapping: {
    flex: 1,
    minHeight: 0,
    borderRadius: 0,
    borderWidth: 0,
    backgroundColor: '#f2f0eb'
  },
  panelOverlay: {
    flex: 1,
    borderRadius: 0,
    borderWidth: 0,
    minHeight: 0,
    backgroundColor: '#f2f0eb'
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
  webviewMapping: {
    flex: 1,
    minHeight: 0,
    backgroundColor: '#f2f0eb'
  },
  webviewOverlay: {
    flex: 1,
    minHeight: 0,
    backgroundColor: '#f2f0eb'
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
  centerOverlay: {
    flex: 1,
    minHeight: 0
  },
  muted: { ...typography.caption, color: colors.muted, textAlign: 'center' }
});
