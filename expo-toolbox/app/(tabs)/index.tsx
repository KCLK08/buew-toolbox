import { useEffect, useRef } from 'react';
import { Animated, ScrollView, StyleSheet, Text } from 'react-native';
import { useRouter } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { ToolCard } from '../../src/components/ToolCard';
import { ToolboxBackground } from '../../src/components/ToolboxBackground';
import { colors, spacing } from '../../src/constants/theme';
import { TOOLBOX_TOOLS } from '../../src/constants/tools';

export default function ToolboxHomeScreen() {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const fadeIn = useRef(new Animated.Value(0)).current;
  const slideUp = useRef(new Animated.Value(18)).current;

  useEffect(() => {
    Animated.parallel([
      Animated.timing(fadeIn, { toValue: 1, duration: 420, useNativeDriver: true }),
      Animated.timing(slideUp, { toValue: 0, duration: 420, useNativeDriver: true })
    ]).start();
  }, [fadeIn, slideUp]);

  return (
    <ToolboxBackground>
      <ScrollView
        contentContainerStyle={{
          paddingTop: Math.max(spacing.pageTop, insets.top + 16),
          paddingBottom: Math.max(spacing.pageBottom, insets.bottom + spacing.tabBarBody + spacing.lg),
          paddingHorizontal: spacing.pageX,
          gap: spacing.md
        }}
        showsVerticalScrollIndicator={false}
      >
        <Animated.View style={{ opacity: fadeIn, transform: [{ translateY: slideUp }], gap: 8 }}>
          <Text style={styles.title}>BÜW-Toolbox</Text>
          <Text style={styles.sub}>
            Native App mit denselben Werkzeugen wie die PWA — ohne WebView, vollständig offline auf
            dem Gerät.
          </Text>
        </Animated.View>

        <Animated.View style={{ opacity: fadeIn, transform: [{ translateY: slideUp }], gap: spacing.cardGap }}>
          {TOOLBOX_TOOLS.map((tool) => (
            <ToolCard key={tool.id} tool={tool} onPress={() => router.push(tool.tabHref)} />
          ))}
        </Animated.View>
      </ScrollView>
    </ToolboxBackground>
  );
}

const styles = StyleSheet.create({
  title: {
    color: colors.ink,
    fontFamily: 'SpaceGrotesk_700Bold',
    fontSize: 34,
    letterSpacing: -0.6
  },
  sub: {
    color: colors.muted,
    fontSize: 16,
    lineHeight: 24,
    fontFamily: 'SpaceGrotesk_400Regular'
  }
});
