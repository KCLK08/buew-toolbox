import { Stack } from 'expo-router';

export default function BautagebuchLayout() {
  return (
    <Stack screenOptions={{ headerShown: false, contentStyle: { backgroundColor: 'transparent' } }}>
      <Stack.Screen name="(tabs)" />
      <Stack.Screen name="run/[id]" />
      <Stack.Screen name="setup" />
    </Stack>
  );
}
