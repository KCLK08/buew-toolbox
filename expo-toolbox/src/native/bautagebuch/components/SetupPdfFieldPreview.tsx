import { useEffect, useMemo, useRef, useState } from 'react';
import { ActivityIndicator, Dimensions, StyleSheet, Text, View } from 'react-native';
import * as FileSystem from 'expo-file-system/legacy';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { WebView, type WebViewMessageEvent } from 'react-native-webview';

import { PrimaryButton } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';
import {
  buildFieldPreviewHtml,
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
  const insets = useSafeAreaInsets();
  const pinned = variant === 'pinned';
  const mapping = variant === 'mapping';
  const overlay = variant === 'overlay';
  const highQuality = overlay || mapping;
  const highlightActive = emphasizeActiveHighlight || mapping || overlay;
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
    void FileSystem.readAsStringAsync(pdfPath, {
      encoding: FileSystem.EncodingType.Base64
    })
      .then((base64) => {
        if (cancelled) return;
        setHtml(
          buildFieldPreviewHtml({
            base64,
            highlights,
            mode: htmlMode,
            highlightActive,
            highQuality
          })
        );
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
  }, [pdfPath, highlights, htmlMode, highlightActive, highQuality]);

  useEffect(() => {
    if (!previewState.ready || useFallback) return;
    webViewRef.current?.postMessage(
      JSON.stringify({
        type: 'setActive',
        fieldId: activeFieldId || '',
        page: Number(activeFieldPage || previewState.page || 1),
        overlayPlacement
      })
    );
  }, [activeFieldId, activeFieldPage, overlayPlacement, previewState.ready, useFallback]);

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

  const banner = mapping
    ? activeFieldLabel || 'Feld zuordnen'
    : activeFieldLabel
      ? `Aktiv: ${activeFieldLabel}${activeFieldPage ? ` · Seite ${activeFieldPage}` : ''}`
      : pinned
        ? 'Feld oder Spalte antippen — Markierung in der PDF zeigt die Position.'
        : 'Feld in der Liste antippen, um die zugehörige PDF-Seite zu sehen.';

  return (
    <View
      style={[
        styles.root,
        pinned ? styles.rootPinned : null,
        mapping ? styles.rootMapping : null,
        overlay ? [styles.rootOverlay, { paddingBottom: Math.max(insets.bottom, 0) }] : null
      ]}
    >
      {!mapping ? (
        <Text style={[styles.banner, pinned ? styles.bannerPinned : null]} numberOfLines={2}>
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
          onError={() => setUseFallback(true)}
          onHttpError={() => setUseFallback(true)}
        />
      </View>
      {!mapping ? (
        <View style={[styles.controls, overlay ? { paddingBottom: Math.max(insets.bottom, spacing.xs) } : null]}>
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
      {highlights.length === 0 && !pinned && !mapping ? (
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
    gap: spacing.xs,
    minHeight: 0
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
  panelMapping: {
    flex: 1,
    minHeight: 0,
    borderRadius: 0,
    borderWidth: 0,
    backgroundColor: '#f2f0eb'
  },
  panelOverlay: {
    flex: 1,
    borderRadius: 10,
    minHeight: 0
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
    minHeight: 0
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
