import { useEffect, useRef } from 'react';
import { Animated, ScrollView, StyleSheet, useWindowDimensions, View } from 'react-native';
import { useRouter } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { ToolCard } from '../../src/components/ToolCard';
import { HomeHeader } from '../../src/components/toolbox/HomeHeader';
import { homeCardsGap, resolveHomeLayoutTier } from '../../src/components/toolbox/homeLayout';
import { ToolboxBackground } from '../../src/components/ToolboxBackground';
import { spacing } from '../../src/constants/theme';
import { TOOLBOX_TOOLS } from '../../src/constants/tools';
import { systemBottomInset } from '../../src/navigation/systemInsets';

export default function ToolboxHomeScreen() {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const { height: windowHeight } = useWindowDimensions();
  const fadeIn = useRef(new Animated.Value(0)).current;
  const slideUp = useRef(new Animated.Value(16)).current;

  const topPad = Math.max(spacing.sm, insets.top + spacing.sm);
  const bottomPad = systemBottomInset(insets) + spacing.lg;
  const contentMinHeight = windowHeight - topPad - bottomPad;
  const tier = resolveHomeLayoutTier(contentMinHeight);
  const cardsGap = homeCardsGap(tier);

  useEffect(() => {
    Animated.parallel([
      Animated.timing(fadeIn, { toValue: 1, duration: 480, useNativeDriver: true }),
      Animated.timing(slideUp, { toValue: 0, duration: 480, useNativeDriver: true })
    ]).start();
  }, [fadeIn, slideUp]);

  return (
    <ToolboxBackground>
      <ScrollView
        style={styles.scroll}
        contentContainerStyle={[
          styles.scrollContent,
          {
            minHeight: contentMinHeight,
            paddingTop: topPad,
            paddingBottom: bottomPad
          }
        ]}
        showsVerticalScrollIndicator={false}
        bounces={false}
        overScrollMode="never"
      >
        <View style={[styles.page, { minHeight: contentMinHeight }]}>
          <Animated.View
            style={[styles.headerWrap, { opacity: fadeIn, transform: [{ translateY: slideUp }] }]}
          >
            <HomeHeader tier={tier} />
          </Animated.View>

          <Animated.View
            style={[
              styles.cards,
              { gap: cardsGap, opacity: fadeIn, transform: [{ translateY: slideUp }] }
            ]}
          >
            {TOOLBOX_TOOLS.map((tool) => (
              <ToolCard
                key={tool.id}
                tool={tool}
                tier={tier}
                onPress={() => router.push(tool.tabHref)}
              />
            ))}
          </Animated.View>
        </View>
      </ScrollView>
    </ToolboxBackground>
  );
}

const styles = StyleSheet.create({
  scroll: {
    flex: 1
  },
  scrollContent: {
    flexGrow: 1,
    paddingHorizontal: spacing.pageX
  },
  page: {
    flexGrow: 1,
    minHeight: 0
  },
  headerWrap: {
    flexShrink: 0
  },
  cards: {
    flex: 1,
    minHeight: 0,
    justifyContent: 'flex-start',
    paddingTop: spacing.xxs
  }
});
