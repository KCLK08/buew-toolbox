import { useEffect, useRef } from 'react';
import { Animated, useWindowDimensions, View } from 'react-native';
import { useRouter } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { ToolCard } from '../../src/components/ToolCard';
import { HomeHeader } from '../../src/components/toolbox/HomeHeader';
import { ToolboxBackground } from '../../src/components/ToolboxBackground';
import { spacing } from '../../src/constants/theme';
import { TOOLBOX_TOOLS } from '../../src/constants/tools';
import { systemBottomInset } from '../../src/navigation/systemInsets';

/** Below this height the home layout compacts cards to stay on one screen. */
const COMPACT_HOME_HEIGHT = 760;

export default function ToolboxHomeScreen() {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const { height: windowHeight } = useWindowDimensions();
  const compact = windowHeight < COMPACT_HOME_HEIGHT;
  const fadeIn = useRef(new Animated.Value(0)).current;
  const slideUp = useRef(new Animated.Value(16)).current;

  const topPad = Math.max(spacing.sm, insets.top + spacing.sm);
  const bottomPad = systemBottomInset(insets) + spacing.md;

  useEffect(() => {
    Animated.parallel([
      Animated.timing(fadeIn, { toValue: 1, duration: 480, useNativeDriver: true }),
      Animated.timing(slideUp, { toValue: 0, duration: 480, useNativeDriver: true })
    ]).start();
  }, [fadeIn, slideUp]);

  return (
    <ToolboxBackground>
      <View
        style={{
          flex: 1,
          paddingTop: topPad,
          paddingBottom: bottomPad,
          paddingHorizontal: spacing.pageX
        }}
      >
        <Animated.View style={{ opacity: fadeIn, transform: [{ translateY: slideUp }] }}>
          <HomeHeader compact={compact} />
        </Animated.View>

        <Animated.View
          style={{
            flex: 1,
            minHeight: 0,
            justifyContent: 'center',
            gap: compact ? spacing.sm : spacing.md,
            opacity: fadeIn,
            transform: [{ translateY: slideUp }]
          }}
        >
          {TOOLBOX_TOOLS.map((tool) => (
            <ToolCard
              key={tool.id}
              tool={tool}
              compact={compact}
              onPress={() => router.push(tool.tabHref)}
            />
          ))}
        </Animated.View>
      </View>
    </ToolboxBackground>
  );
}
