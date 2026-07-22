import { Stack } from 'expo-router';

export default function SetupLayout() {
  return (
    <Stack screenOptions={{ headerShown: false }}>
      <Stack.Screen name="index" />
      <Stack.Screen name="[templateId]/mapping" />
      <Stack.Screen name="[templateId]/fields" />
    </Stack>
  );
}
