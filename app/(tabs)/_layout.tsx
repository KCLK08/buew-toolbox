import { Tabs } from 'expo-router';
import { Platform } from 'react-native';

import { colors } from '@/theme/colors';

export default function TabLayout() {
  return (
    <Tabs
      screenOptions={{
        tabBarActiveTintColor: colors.primary,
        headerStyle: { backgroundColor: colors.primary },
        headerTintColor: '#fff',
        headerTitleStyle: { fontWeight: '700' },
      }}
    >
      <Tabs.Screen
        name="index"
        options={{
          title: 'Bautagebücher',
          tabBarLabel: 'BTB',
        }}
      />
      <Tabs.Screen
        name="templates"
        options={{
          title: 'Vorlagen',
          tabBarLabel: 'Vorlagen',
        }}
      />
    </Tabs>
  );
}
