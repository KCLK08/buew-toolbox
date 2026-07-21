import { useCallback, useRef, useState } from 'react';
import { ActivityIndicator, Pressable, StyleSheet, Text, View } from 'react-native';
import { WebView, type WebViewNavigation } from 'react-native-webview';
import { useRouter } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../constants/theme';
import { toolWebUrl } from '../lib/config';

type ToolWebScreenProps = {
  title: string;
  webPath: string;
  mode?: 'tab' | 'stack';
};

export function ToolWebScreen({ title, webPath, mode = 'stack' }: ToolWebScreenProps) {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const webRef = useRef<WebView>(null);
  const [canGoBack, setCanGoBack] = useState(false);
  const [loading, setLoading] = useState(true);
  const uri = toolWebUrl(webPath);
  const showBack = mode === 'stack';

  const onNavigationStateChange = useCallback((state: WebViewNavigation) => {
    setCanGoBack(state.canGoBack);
  }, []);

  const handleBack = () => {
    if (mode === 'stack') {
      router.back();
      return;
    }
    if (canGoBack) {
      webRef.current?.goBack();
    }
  };

  return (
    <View style={[styles.root, { paddingTop: insets.top }]}>
      <View style={styles.toolbar}>
        {showBack || canGoBack ? (
          <Pressable
            accessibilityRole="button"
            accessibilityLabel="Zurück"
            onPress={handleBack}
            style={styles.backButton}
            hitSlop={8}
          >
            <Text style={styles.backLabel}>‹ Zurück</Text>
          </Pressable>
        ) : (
          <View style={styles.backPlaceholder} />
        )}
        <Text style={styles.title} numberOfLines={1}>
          {title}
        </Text>
        <View style={styles.backPlaceholder} />
      </View>
      <WebView
        ref={webRef}
        source={{ uri }}
        style={styles.webview}
        startInLoadingState
        onLoadStart={() => setLoading(true)}
        onLoadEnd={() => setLoading(false)}
        onNavigationStateChange={onNavigationStateChange}
        javaScriptEnabled
        domStorageEnabled
        allowsInlineMediaPlayback
        mediaPlaybackRequiresUserAction={false}
        allowsBackForwardNavigationGestures
        setSupportMultipleWindows={false}
        originWhitelist={['*']}
        mixedContentMode="compatibility"
        sharedCookiesEnabled
        thirdPartyCookiesEnabled
        cacheEnabled
        pullToRefreshEnabled
      />
      {loading ? (
        <View style={styles.loading} pointerEvents="none">
          <ActivityIndicator color={colors.accent} size="large" />
          <Text style={styles.loadingLabel}>Werkzeug wird geladen…</Text>
        </View>
      ) : null}
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg
  },
  toolbar: {
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.pageX,
    flexDirection: 'row',
    alignItems: 'center',
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panel
  },
  backButton: {
    minHeight: spacing.touchMin,
    justifyContent: 'center',
    paddingRight: spacing.sm,
    minWidth: 72
  },
  backPlaceholder: {
    minWidth: 72
  },
  backLabel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  title: {
    flex: 1,
    textAlign: 'center',
    ...typography.subtitle,
    color: colors.ink
  },
  webview: {
    flex: 1,
    backgroundColor: colors.bg
  },
  loading: {
    ...StyleSheet.absoluteFill,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.bg,
    gap: spacing.sm
  },
  loadingLabel: {
    ...typography.caption,
    color: colors.muted
  }
});
