import { useEffect, useRef, useState } from 'react';
import {
  Animated,
  ScrollView,
  StyleSheet,
  Text,
  useWindowDimensions,
  View
} from 'react-native';
import { useRouter } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { OfflineStatusBanner } from '../src/components/OfflineStatusBanner';
import { ToolCard } from '../src/components/ToolCard';
import { ToolboxBackground } from '../src/components/ToolboxBackground';
import { colors, spacing } from '../src/constants/theme';
import { TOOLBOX_TOOLS } from '../src/constants/tools';
import { useOfflineBootstrap } from '../src/hooks/useOfflineBootstrap';

export default function ToolboxHomeScreen() {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const { width } = useWindowDimensions();
  const offline = useOfflineBootstrap();
  const [bannerDismissed, setBannerDismissed] = useState(false);
  const fadeIn = useRef(new Animated.Value(0)).current;
  const slideUp = useRef(new Animated.Value(18)).current;

  useEffect(() => {
    Animated.parallel([
      Animated.timing(fadeIn, {
        toValue: 1,
        duration: 420,
        useNativeDriver: true
      }),
      Animated.timing(slideUp, {
        toValue: 0,
        duration: 420,
        useNativeDriver: true
      })
    ]).start();
  }, [fadeIn, slideUp]);

  const isWide = width >= 840;
  const titleSize = Math.min(52, Math.max(32, width * 0.08));

  return (
    <ToolboxBackground>
      <ScrollView
        contentContainerStyle={[
          styles.page,
          {
            paddingTop: Math.max(spacing.pageTop, insets.top + 24),
            paddingBottom: Math.max(spacing.pageBottom, insets.bottom + 48)
          }
        ]}
        showsVerticalScrollIndicator={false}
      >
        <Animated.View
          style={[
            styles.hero,
            isWide ? styles.heroWide : null,
            {
              opacity: fadeIn,
              transform: [{ translateY: slideUp }]
            }
          ]}
        >
          <View style={styles.heroCopy}>
            <Text style={[styles.title, { fontSize: titleSize }]}>BÜW-Toolbox</Text>
            <Text style={styles.sub}>
              Zentrale Übersicht für digitale Baustellen‑Workflows und Dokumentation.
            </Text>
            <Text style={styles.offlineHint}>Offline · Daten bleiben lokal auf diesem Gerät</Text>
          </View>
        </Animated.View>

        {!bannerDismissed ? (
          <OfflineStatusBanner
            report={offline.report}
            error={offline.error}
            onDismiss={() => setBannerDismissed(true)}
          />
        ) : null}

        <Animated.View
          style={[
            styles.grid,
            {
              opacity: fadeIn,
              transform: [{ translateY: slideUp }]
            }
          ]}
        >
          {TOOLBOX_TOOLS.map((tool) => (
            <ToolCard
              key={tool.id}
              tool={tool}
              onPress={() => router.push(tool.href)}
            />
          ))}
        </Animated.View>
      </ScrollView>
    </ToolboxBackground>
  );
}

const styles = StyleSheet.create({
  page: {
    width: '100%',
    maxWidth: 1100,
    alignSelf: 'center',
    paddingHorizontal: spacing.pageX
  },
  hero: {
    marginBottom: 36
  },
  heroWide: {
    maxWidth: '100%'
  },
  heroCopy: {
    maxWidth: 640
  },
  title: {
    color: colors.ink,
    fontFamily: 'SpaceGrotesk_700Bold',
    letterSpacing: -0.6,
    marginBottom: 12
  },
  sub: {
    color: colors.muted,
    fontSize: 16,
    lineHeight: 24,
    fontFamily: 'SpaceGrotesk_400Regular',
    marginBottom: 12
  },
  offlineHint: {
    color: colors.accent2,
    fontSize: 13,
    fontFamily: 'SpaceGrotesk_400Regular',
    opacity: 0.85
  },
  grid: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.cardGap,
    alignItems: 'stretch'
  }
});
