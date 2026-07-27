import { useEffect, useRef } from 'react';
import { Animated, ScrollView, View } from 'react-native';
import { useRouter } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { ToolCard } from '../../src/components/ToolCard';
import { HomeHeader } from '../../src/components/toolbox/HomeHeader';
import { ToolboxBackground } from '../../src/components/ToolboxBackground';
import { spacing } from '../../src/constants/theme';
import { TOOLBOX_TOOLS } from '../../src/constants/tools';
import { tabBarReservedHeight } from '../../src/navigation/appTabBar';

export default function ToolboxHomeScreen() {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const fadeIn = useRef(new Animated.Value(0)).current;
  const slideUp = useRef(new Animated.Value(16)).current;

  useEffect(() => {
    Animated.parallel([
      Animated.timing(fadeIn, { toValue: 1, duration: 480, useNativeDriver: true }),
      Animated.timing(slideUp, { toValue: 0, duration: 480, useNativeDriver: true })
    ]).start();
  }, [fadeIn, slideUp]);

  return (
    <ToolboxBackground>
      <ScrollView
        contentContainerStyle={{
          paddingTop: Math.max(spacing.pageTop, insets.top + 16),
          paddingBottom: Math.max(spacing.pageBottom, tabBarReservedHeight(insets) + spacing.xl),
          paddingHorizontal: spacing.pageX,
          gap: spacing.lg
        }}
        showsVerticalScrollIndicator={false}
      >
        <Animated.View style={{ opacity: fadeIn, transform: [{ translateY: slideUp }] }}>
          <HomeHeader />
        </Animated.View>

        <Animated.View style={{ opacity: fadeIn, transform: [{ translateY: slideUp }], gap: spacing.md }}>
          {TOOLBOX_TOOLS.map((tool) => (
            <ToolCard key={tool.id} tool={tool} onPress={() => router.push(tool.tabHref)} />
          ))}
        </Animated.View>
      </ScrollView>
    </ToolboxBackground>
  );
}
