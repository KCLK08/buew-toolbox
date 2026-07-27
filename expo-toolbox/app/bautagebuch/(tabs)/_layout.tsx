import { Tabs } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { createAppTabScreenOptions } from '../../../src/navigation/appTabBar';
import { TabIcon } from '../../../src/navigation/TabIcon';

export default function BautagebuchTabsLayout() {
  const insets = useSafeAreaInsets();

  return (
    <Tabs backBehavior="none" screenOptions={createAppTabScreenOptions(insets)}>
      <Tabs.Screen
        name="btbs"
        options={{
          title: 'BTBs',
          tabBarIcon: ({ focused }) => <TabIcon name="format-list-bulleted" focused={focused} />
        }}
      />
      <Tabs.Screen
        name="index"
        options={{
          title: 'Home',
          tabBarIcon: ({ focused }) => <TabIcon name="home-outline" focused={focused} />
        }}
      />
      <Tabs.Screen
        name="config"
        options={{
          title: 'Setup',
          tabBarIcon: ({ focused }) => <TabIcon name="cog-outline" focused={focused} />
        }}
      />
    </Tabs>
  );
}
