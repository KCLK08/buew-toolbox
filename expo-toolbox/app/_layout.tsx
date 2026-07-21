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

import { colors } from '../src/constants/theme';
import { initBautagebuchDatabase } from '../src/native/bautagebuch/db/database';
import { initSiteReportDatabase } from '../src/native/sitereport/db/database';

export default function RootLayout() {
  const [fontsLoaded] = useFonts({
    SpaceGrotesk_400Regular,
    SpaceGrotesk_600SemiBold,
    SpaceGrotesk_700Bold
  });
  const [ready, setReady] = useState(false);

  useEffect(() => {
    Promise.all([initBautagebuchDatabase(), initSiteReportDatabase()])
      .then(() => setReady(true))
      .catch(() => setReady(true));
  }, []);

  if (!fontsLoaded || !ready) {
    return (
      <View style={{ flex: 1, alignItems: 'center', justifyContent: 'center', backgroundColor: colors.bg, gap: 12 }}>
        <ActivityIndicator color={colors.accent} size="large" />
        <Text style={{ color: colors.muted, fontFamily: 'SpaceGrotesk_400Regular' }}>App wird geladen…</Text>
      </View>
    );
  }

  return (
    <GestureHandlerRootView style={{ flex: 1 }}>
      <StatusBar style="dark" />
      <Stack screenOptions={{ headerShown: false, contentStyle: { backgroundColor: colors.bg } }}>
        <Stack.Screen name="(tabs)" />
        <Stack.Screen name="bautagebuch/run/[id]" />
        <Stack.Screen name="sitereport/protocol/[id]" />
      </Stack>
    </GestureHandlerRootView>
  );
}
