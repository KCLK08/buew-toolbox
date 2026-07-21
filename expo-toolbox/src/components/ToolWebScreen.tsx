import { ActivityIndicator, Pressable, StyleSheet, Text, View } from 'react-native';
import { WebView } from 'react-native-webview';
import { useRouter } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors } from '../constants/theme';
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
          accessibilityLabel="Zurück zur Toolbox"
          onPress={() => router.back()}
          style={styles.backButton}
        >
          <Text style={styles.backLabel}>← Toolbox</Text>
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
    minHeight: 52,
    paddingHorizontal: 12,
    flexDirection: 'row',
    alignItems: 'center',
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panel
  },
  backButton: {
    paddingVertical: 8,
    paddingHorizontal: 8
  },
  backLabel: {
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 15
  },
  title: {
    flex: 1,
    textAlign: 'center',
    color: colors.ink,
    fontFamily: 'SpaceGrotesk_700Bold',
    fontSize: 16
  },
  spacer: {
    width: 88
  },
  webview: {
    flex: 1,
    backgroundColor: colors.bg
  },
  loading: {
    ...StyleSheet.absoluteFill,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.bg
  }
});
