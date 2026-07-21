import { useEffect, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import * as FileSystem from 'expo-file-system/legacy';
import { WebView } from 'react-native-webview';

import { colors, spacing, typography } from '../../../constants/theme';

type Props = {
  pdfPath: string | null;
  loading?: boolean;
  error?: string | null;
};

export function PdfPreviewPanel({ pdfPath, loading = false, error = null }: Props) {
  const [html, setHtml] = useState<string | null>(null);
  const [readBusy, setReadBusy] = useState(false);

  useEffect(() => {
    let cancelled = false;
    if (!pdfPath) {
      setHtml(null);
      return undefined;
    }

    setReadBusy(true);
    void FileSystem.readAsStringAsync(pdfPath, {
      encoding: FileSystem.EncodingType.Base64
    })
      .then((base64) => {
        if (cancelled) return;
        setHtml(`<!DOCTYPE html>
<html>
  <head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0" />
    <style>
      html, body { margin: 0; padding: 0; height: 100%; background: #f2f0eb; }
      embed { width: 100%; height: 100%; border: 0; }
    </style>
  </head>
  <body>
    <embed src="data:application/pdf;base64,${base64}" type="application/pdf" />
  </body>
</html>`);
      })
      .catch(() => {
        if (!cancelled) setHtml(null);
      })
      .finally(() => {
        if (!cancelled) setReadBusy(false);
      });

    return () => {
      cancelled = true;
    };
  }, [pdfPath]);

  if (loading || readBusy) {
    return (
      <View style={styles.center}>
        <ActivityIndicator color={colors.accent} />
        <Text style={styles.muted}>PDF-Vorschau wird aktualisiert…</Text>
      </View>
    );
  }

  if (error) {
    return (
      <View style={styles.center}>
        <Text style={styles.error}>{error}</Text>
      </View>
    );
  }

  if (!html) {
    return (
      <View style={styles.center}>
        <Text style={styles.muted}>Noch keine Vorschau verfügbar.</Text>
      </View>
    );
  }

  return (
    <View style={styles.panel}>
      <WebView
        source={{ html }}
        style={styles.webview}
        originWhitelist={['*']}
        scrollEnabled
      />
    </View>
  );
}

const styles = StyleSheet.create({
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
  center: {
    minHeight: 200,
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm,
    padding: spacing.md
  },
  muted: { ...typography.caption, color: colors.muted, textAlign: 'center' },
  error: { ...typography.body, color: colors.danger, textAlign: 'center' }
});
