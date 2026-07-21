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
import { useOfflineBootstrap } from '../src/hooks/useOfflineBootstrap';

export default function RootLayout() {
  const [fontsLoaded] = useFonts({
    SpaceGrotesk_400Regular,
    SpaceGrotesk_600SemiBold,
    SpaceGrotesk_700Bold
  });
  const offline = useOfflineBootstrap();

  if (!fontsLoaded || !offline.ready) {
    return (
      <View
        style={{
          flex: 1,
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: colors.bg,
          gap: 12
        }}
      >
        <ActivityIndicator color={colors.accent} size="large" />
        <Text style={{ color: colors.muted, fontFamily: 'SpaceGrotesk_400Regular' }}>
          Offline-Daten werden geprüft…
        </Text>
      </View>
    );
  }

  return (
    <GestureHandlerRootView style={{ flex: 1 }}>
      <StatusBar style="dark" />
      <Stack
        screenOptions={{
          headerShown: false,
          contentStyle: { backgroundColor: colors.bg },
          animation: 'slide_from_right'
        }}
      >
        <Stack.Screen name="(tabs)" />
        <Stack.Screen name="project/new" options={{ presentation: 'modal' }} />
        <Stack.Screen name="project/[id]" />
        <Stack.Screen name="diary/new" options={{ presentation: 'modal' }} />
        <Stack.Screen name="diary/[id]" />
        <Stack.Screen name="defect/new" options={{ presentation: 'modal' }} />
        <Stack.Screen name="defect/[id]" />
        <Stack.Screen name="sitereport" />
        <Stack.Screen name="bautagebuch" />
      </Stack>
    </GestureHandlerRootView>
  );
}
