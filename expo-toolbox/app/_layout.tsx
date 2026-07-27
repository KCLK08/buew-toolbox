import 'react-native-gesture-handler';
import { useCallback, useEffect, useState } from 'react';
import { Stack } from 'expo-router';
import { StatusBar } from 'expo-status-bar';
import {
  useFonts,
  SpaceGrotesk_400Regular,
  SpaceGrotesk_600SemiBold,
  SpaceGrotesk_700Bold
} from '@expo-google-fonts/space-grotesk';
import { GestureHandlerRootView } from 'react-native-gesture-handler';
import { SafeAreaProvider } from 'react-native-safe-area-context';

import { AppLoadingScreen } from '../src/components/AppLoadingScreen';
import { colors } from '../src/constants/theme';
import { ToastProvider } from '../src/contexts/ToastContext';
import { useOfflineBootstrap } from '../src/hooks/useOfflineBootstrap';
import { AndroidHardwareBack } from '../src/navigation/useAndroidHardwareBack';
import { initBautagebuchDatabase } from '../src/native/bautagebuch/db/database';
import { initSiteReportDatabase } from '../src/native/sitereport/db/database';

export default function RootLayout() {
  return (
    <SafeAreaProvider>
      <RootLayoutContent />
    </SafeAreaProvider>
  );
}

function RootLayoutContent() {
  const [fontsLoaded] = useFonts({
    SpaceGrotesk_400Regular,
    SpaceGrotesk_600SemiBold,
    SpaceGrotesk_700Bold
  });
  const [nativeReady, setNativeReady] = useState(false);
  const [showSplash, setShowSplash] = useState(true);
  useOfflineBootstrap();

  useEffect(() => {
    Promise.all([initBautagebuchDatabase(), initSiteReportDatabase()])
      .then(() => setNativeReady(true))
      .catch(() => setNativeReady(true));
  }, []);

  const appReady = fontsLoaded && nativeReady;
  const handleSplashFinish = useCallback(() => setShowSplash(false), []);

  if (!appReady) {
    return (
      <GestureHandlerRootView style={{ flex: 1, backgroundColor: colors.bg }}>
        <AppLoadingScreen onFinish={() => undefined} minDurationMs={60000} />
      </GestureHandlerRootView>
    );
  }

  if (showSplash) {
    return (
      <GestureHandlerRootView style={{ flex: 1 }}>
        <AppLoadingScreen onFinish={handleSplashFinish} />
      </GestureHandlerRootView>
    );
  }

  return (
    <GestureHandlerRootView style={{ flex: 1 }}>
      <ToastProvider>
        <StatusBar style="dark" />
        <AndroidHardwareBack />
        <Stack screenOptions={{ headerShown: false, contentStyle: { backgroundColor: colors.bg } }}>
          <Stack.Screen name="(tabs)" />
          <Stack.Screen name="bautagebuch" />
          <Stack.Screen name="sitereport" />
        </Stack>
      </ToastProvider>
    </GestureHandlerRootView>
  );
}
