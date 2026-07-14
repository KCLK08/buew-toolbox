import { Stack } from 'expo-router';
import { StatusBar } from 'expo-status-bar';

export default function RootLayout() {
  return (
    <>
      <Stack
        screenOptions={{
          headerStyle: { backgroundColor: '#12534b' },
          headerTintColor: '#fff',
          headerTitleStyle: { fontWeight: '700' },
        }}
      >
        <Stack.Screen name="(tabs)" options={{ headerShown: false }} />
        <Stack.Screen name="run/[id]" options={{ title: 'Bautagebuch' }} />
        <Stack.Screen name="setup/[templateId]" options={{ title: 'Vorlage Setup' }} />
      </Stack>
      <StatusBar style="light" />
    </>
  );
}
