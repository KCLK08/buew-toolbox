import { ActivityIndicator, Pressable, StyleSheet, Text, View } from 'react-native';
import { WebView } from 'react-native-webview';
import { useRouter } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../constants/theme';
import { toolWebUrl } from '../lib/config';

type ToolWebScreenProps = {
  title: string;
  webPath: string;
};

export function ToolWebScreen({ title, webPath }: ToolWebScreenProps) {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const uri = toolWebUrl(webPath);

  return (
    <View style={[styles.root, { paddingTop: insets.top }]}>
      <View style={styles.toolbar}>
        <Pressable
          accessibilityRole="button"
          accessibilityLabel="Zurück"
          onPress={() => router.back()}
          style={styles.backButton}
          hitSlop={8}
        >
          <Text style={styles.backLabel}>‹ Zurück</Text>
        </Pressable>
        <Text style={styles.title} numberOfLines={1}>
          {title}
        </Text>
        <View style={styles.spacer} />
      </View>
      <WebView
        source={{ uri }}
        style={styles.webview}
        startInLoadingState
        renderLoading={() => (
          <View style={styles.loading}>
            <ActivityIndicator color={colors.accent} size="large" />
          </View>
        )}
        allowsBackForwardNavigationGestures
        setSupportMultipleWindows={false}
      />
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg
  },
  toolbar: {
    minHeight: spacing.touchMin + 8,
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
    paddingRight: spacing.sm
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
  spacer: {
    width: 72
  },
  webview: {
    flex: 1,
    backgroundColor: colors.bg
  },
  loading: {
    ...StyleSheet.absoluteFillObject,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.bg
  }
});
