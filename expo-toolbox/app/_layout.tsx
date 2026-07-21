import { useEffect, useState } from 'react';
import { Stack } from 'expo-router';
import { StatusBar } from 'expo-status-bar';
import {
  useFonts,
  SpaceGrotesk_400Regular,
  SpaceGrotesk_600SemiBold,
  SpaceGrotesk_700Bold
} from '@expo-google-fonts/space-grotesk';
import { ActivityIndicator, Text, View } from 'react-native';
import { GestureHandlerRootView } from 'react-native-gesture-handler';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { OfflineStatusBanner } from '../src/components/OfflineStatusBanner';
import { colors } from '../src/constants/theme';
import { ToastProvider } from '../src/contexts/ToastContext';
import { useOfflineBootstrap } from '../src/hooks/useOfflineBootstrap';
import { initBautagebuchDatabase } from '../src/native/bautagebuch/db/database';
import { initSiteReportDatabase } from '../src/native/sitereport/db/database';

export default function RootLayout() {
  const insets = useSafeAreaInsets();
  const [fontsLoaded] = useFonts({
    SpaceGrotesk_400Regular,
    SpaceGrotesk_600SemiBold,
    SpaceGrotesk_700Bold
  });
  const [nativeReady, setNativeReady] = useState(false);
  const [bannerDismissed, setBannerDismissed] = useState(false);
  const { ready: offlineReady, report, error, restoreBusy, acceptRestore, rejectRestore } =
    useOfflineBootstrap();
  const showBanner = report?.pendingRestore || report?.restoredFromBackup || (!bannerDismissed && (error || (report && !report.ok)));

  useEffect(() => {
    Promise.all([initBautagebuchDatabase(), initSiteReportDatabase()])
      .then(() => setNativeReady(true))
      .catch(() => setNativeReady(true));
  }, []);

  const appReady = fontsLoaded && nativeReady && offlineReady;

  if (!appReady) {
    return (
      <View style={{ flex: 1, alignItems: 'center', justifyContent: 'center', backgroundColor: colors.bg, gap: 12 }}>
        <ActivityIndicator color={colors.accent} size="large" />
        <Text style={{ color: colors.muted, fontFamily: 'SpaceGrotesk_400Regular' }}>App wird geladen…</Text>
      </View>
    );
  }

  return (
    <GestureHandlerRootView style={{ flex: 1 }}>
      <ToastProvider>
      <StatusBar style="dark" />
      <View style={{ flex: 1, paddingTop: insets.top }}>
        <View style={{ paddingHorizontal: 16 }}>
          {showBanner ? (
            <OfflineStatusBanner
              report={report}
              error={error}
              restoreBusy={restoreBusy}
              onAcceptRestore={() => {
                void acceptRestore().then(() => {
                  void Promise.all([initBautagebuchDatabase(), initSiteReportDatabase()]);
                });
              }}
              onRejectRestore={rejectRestore}
              onDismiss={() => setBannerDismissed(true)}
            />
          ) : null}
        </View>
        <Stack screenOptions={{ headerShown: false, contentStyle: { backgroundColor: colors.bg } }}>
          <Stack.Screen name="(tabs)" />
          <Stack.Screen name="bautagebuch/run/[id]" />
          <Stack.Screen name="bautagebuch/setup" />
          <Stack.Screen name="sitereport/new-protocol" />
          <Stack.Screen name="sitereport/protocol/[id]" />
          <Stack.Screen name="sitereport/protocol/[id]/edit" />
          <Stack.Screen name="sitereport/protocol/[id]/wizard" />
          <Stack.Screen name="sitereport/protocols/index" />
          <Stack.Screen name="sitereport/exports/index" />
          <Stack.Screen name="sitereport/format-builder" />
        </Stack>
      </View>
      </ToastProvider>
    </GestureHandlerRootView>
  );
}
