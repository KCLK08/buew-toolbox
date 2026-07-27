import { Stack } from 'expo-router';

export default function SiteReportLayout() {
  return (
    <Stack screenOptions={{ headerShown: false }}>
      <Stack.Screen name="(tabs)" />
      <Stack.Screen name="new-protocol" />
      <Stack.Screen name="protocol/[id]" />
      <Stack.Screen name="format-builder" />
    </Stack>
  );
}
