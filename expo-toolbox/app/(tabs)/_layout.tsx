import { Tabs } from 'expo-router';
import { StyleSheet, Text } from 'react-native';

import { colors, spacing, typography } from '../../src/constants/theme';

function TabIcon({ label, focused }: { label: string; focused: boolean }) {
  return (
    <Text style={[styles.icon, focused ? styles.iconFocused : null]} accessibilityElementsHidden>
      {label}
    </Text>
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
      <Tabs.Screen
        name="index"
        options={{
          title: 'Home',
          tabBarIcon: ({ focused }) => <TabIcon label="⌂" focused={focused} />
        }}
      />
      <Tabs.Screen
        name="projects"
        options={{
          title: 'Projekte',
          tabBarIcon: ({ focused }) => <TabIcon label="▤" focused={focused} />
        }}
      />
      <Tabs.Screen
        name="diary"
        options={{
          title: 'Tagebuch',
          tabBarIcon: ({ focused }) => <TabIcon label="✎" focused={focused} />
        }}
      />
      <Tabs.Screen
        name="defects"
        options={{
          title: 'Mängel',
          tabBarIcon: ({ focused }) => <TabIcon label="!" focused={focused} />
        }}
      />
      <Tabs.Screen
        name="more"
        options={{
          title: 'Mehr',
          tabBarIcon: ({ focused }) => <TabIcon label="⋯" focused={focused} />
        }}
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
  tabItem: {
    minHeight: spacing.touchMin
  },
  tabLabel: {
    ...typography.caption,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 11
  },
  icon: {
    fontSize: 18,
    color: colors.tabInactive
  },
  iconFocused: {
    color: colors.tabActive
  }
});
