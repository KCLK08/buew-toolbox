import { Tabs } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { createAppTabScreenOptions } from '../../../src/navigation/appTabBar';
import { TabIcon } from '../../../src/navigation/TabIcon';

export default function SiteReportTabsLayout() {
  const insets = useSafeAreaInsets();

  return (
    <Tabs backBehavior="none" screenOptions={createAppTabScreenOptions(insets)}>
      <Tabs.Screen
        name="index"
        options={{
          title: 'Home',
          tabBarIcon: ({ focused }) => <TabIcon name="home-outline" focused={focused} />
        }}
      />
      <Tabs.Screen
        name="protocols"
        options={{
          title: 'Protokolle',
          tabBarIcon: ({ focused }) => <TabIcon name="clipboard-text-outline" focused={focused} />
        }}
      />
      <Tabs.Screen
        name="exports"
        options={{
          title: 'Exporte',
          tabBarIcon: ({ focused }) => <TabIcon name="export-variant" focused={focused} />
        }}
      />
    </Tabs>
  );
}
