import { Tabs } from 'expo-router';
import { Image, StyleSheet, Text } from 'react-native';

import { colors, spacing, typography } from '../../src/constants/theme';
import { TOOLBOX_TOOLS } from '../../src/constants/tools';

function TabIcon({ label, focused }: { label: string; focused: boolean }) {
  return (
    <Text style={[styles.icon, focused ? styles.iconFocused : null]} accessibilityElementsHidden>
      {label}
    </Text>
  );
}

function ToolTabIcon({ toolId, focused }: { toolId: 'sitereport' | 'bautagebuch'; focused: boolean }) {
  const tool = TOOLBOX_TOOLS.find((entry) => entry.id === toolId);
  if (!tool) return <TabIcon label="?" focused={focused} />;
  return (
    <Image
      source={tool.icon}
      style={[styles.toolIcon, focused ? styles.toolIconFocused : null]}
      accessibilityLabel={tool.iconAlt}
    />
  );
}

export default function TabsLayout() {
  return (
    <Tabs
      screenOptions={{
        headerShown: false,
        tabBarActiveTintColor: colors.tabActive,
        tabBarInactiveTintColor: colors.tabInactive,
        tabBarLabelStyle: styles.tabLabel,
        tabBarStyle: styles.tabBar,
        tabBarItemStyle: styles.tabItem
      }}
    >
      <Tabs.Screen name="index" options={{ title: 'Home', tabBarIcon: ({ focused }) => <TabIcon label="⌂" focused={focused} /> }} />
      <Tabs.Screen
        name="sitereport"
        options={{ title: 'SiteReport', tabBarIcon: ({ focused }) => <ToolTabIcon toolId="sitereport" focused={focused} /> }}
      />
      <Tabs.Screen
        name="bautagebuch"
        options={{ title: 'Bautagebuch', tabBarIcon: ({ focused }) => <ToolTabIcon toolId="bautagebuch" focused={focused} /> }}
      />
    </Tabs>
  );
}

const styles = StyleSheet.create({
  tabBar: {
    backgroundColor: colors.panel,
    borderTopColor: colors.border,
    borderTopWidth: 1,
    height: 64,
    paddingBottom: 6,
    paddingTop: 6
  },
  tabItem: { minHeight: spacing.touchMin },
  tabLabel: { ...typography.caption, fontFamily: 'SpaceGrotesk_600SemiBold', fontSize: 11 },
  icon: { fontSize: 18, color: colors.tabInactive },
  iconFocused: { color: colors.tabActive },
  toolIcon: { width: 22, height: 22, borderRadius: 6, opacity: 0.72 },
  toolIconFocused: { opacity: 1 }
});
