import { useEffect, useRef, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import * as FileSystem from 'expo-file-system/legacy';
import { WebView, type WebViewMessageEvent } from 'react-native-webview';

import { PrimaryButton } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';
import { buildSimplePdfPreviewHtml } from '../lib/pdf-preview-html';

type Props = {
  pdfPath: string | null;
  loading?: boolean;
  error?: string | null;
};

export function PdfPreviewPanel({ pdfPath, loading = false, error = null }: Props) {
  const webViewRef = useRef<WebView>(null);
  const [html, setHtml] = useState<string | null>(null);
  const [readBusy, setReadBusy] = useState(false);
  const [renderError, setRenderError] = useState<string | null>(null);
  const [previewState, setPreviewState] = useState({
    page: 1,
    pageCount: 1,
    ready: false
  });

  useEffect(() => {
    let cancelled = false;
    if (!pdfPath) {
      return undefined;
    }

    setReadBusy(true);
    setRenderError(null);
    void FileSystem.readAsStringAsync(pdfPath, {
      encoding: FileSystem.EncodingType.Base64
    })
      .then((base64) => {
        if (cancelled) return;
        setHtml(buildSimplePdfPreviewHtml(base64));
        setPreviewState({ page: 1, pageCount: 1, ready: false });
      })
      .catch((readError) => {
        if (!cancelled) {
          setRenderError(readError instanceof Error ? readError.message : 'PDF konnte nicht gelesen werden.');
        }
      })
      .finally(() => {
        if (!cancelled) setReadBusy(false);
      });

    return () => {
      cancelled = true;
    };
  }, [pdfPath]);

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
        setRenderError(payload.error);
        return;
      }
      setRenderError(null);
      setPreviewState({
        page: Number(payload.page || 1),
        pageCount: Number(payload.pageCount || 1),
        ready: Boolean(payload.ready)
      });
    } catch {
      setRenderError('PDF-Vorschau konnte nicht dargestellt werden.');
    }
  };

  const displayError = error || renderError;
  const showInitialLoader = !html && (loading || readBusy || !pdfPath);
  const showUpdateOverlay = Boolean(html) && (loading || readBusy);

  if (displayError && !html) {
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

  if (!html) {
    return (
      <View style={styles.center}>
        <Text style={styles.muted}>Noch keine Live-Vorschau verfügbar.</Text>
      </View>
    );
  }

  return (
    <View style={styles.root}>
      <Text style={styles.banner}>Live-PDF-Vorschau</Text>
      <View style={styles.panel}>
        <WebView
          ref={webViewRef}
          source={{ html }}
          style={styles.webview}
          originWhitelist={['*']}
          scrollEnabled={false}
          bounces={false}
          javaScriptEnabled
          domStorageEnabled
          setBuiltInZoomControls={false}
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
        {displayError ? (
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
  root: { flex: 1, gap: spacing.sm, minHeight: 0 },
  banner: { ...typography.label, color: colors.accent2 },
  panel: {
    flex: 1,
    minHeight: 280,
    borderRadius: 12,
    overflow: 'hidden',
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  webview: {
    flex: 1,
    minHeight: 280,
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
    flex: 1,
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
