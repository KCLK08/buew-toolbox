import { useEffect, useRef, useState } from 'react';
import { Animated, ScrollView, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { ToolCard } from '../../src/components/ToolCard';
import { ToolboxBackground } from '../../src/components/ToolboxBackground';
import { colors, spacing, typography } from '../../src/constants/theme';
import { TOOLBOX_TOOLS } from '../../src/constants/tools';

export default function ToolboxHomeScreen() {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const fadeIn = useRef(new Animated.Value(0)).current;
  const slideUp = useRef(new Animated.Value(18)).current;
  const [hint, setHint] = useState('');

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

  return (
    <ToolboxBackground>
      <ScrollView
        contentContainerStyle={[
          styles.page,
          {
            paddingTop: Math.max(spacing.pageTop, insets.top + 16),
            paddingBottom: Math.max(spacing.pageBottom, insets.bottom + 24)
          }
        ]}
        showsVerticalScrollIndicator={false}
      >
        <Animated.View
          style={[
            styles.hero,
            {
              opacity: fadeIn,
              transform: [{ translateY: slideUp }]
            }
          ]}
        >
          <Text style={styles.title}>BÜW-Toolbox</Text>
          <Text style={styles.sub}>
            Digitale Baustellen-Workflows als App — dieselben Werkzeuge wie in der PWA, mit
            nativer Navigation.
          </Text>
          <Text style={styles.offlineHint}>Offline-fähig · Daten bleiben lokal auf diesem Gerät</Text>
        </Animated.View>

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
              onPress={() => {
                setHint(`Öffne ${tool.title} …`);
                router.push(tool.tabHref);
              }}
            />
          ))}
        </Animated.View>

        {hint ? <Text style={styles.hint}>{hint}</Text> : null}

        <View style={styles.noteCard}>
          <Text style={styles.noteTitle}>Hinweis</Text>
          <Text style={styles.noteCopy}>
            SiteReport und Bautagebuch laufen als eingebettete PWA. Alle Funktionen (PDF-Export,
            Protokolle, Vorlagen, Fotos) sind in den Tabs verfügbar.
          </Text>
        </View>
      </ScrollView>
    </ToolboxBackground>
  );
}

const styles = StyleSheet.create({
  page: {
    width: '100%',
    maxWidth: 1100,
    alignSelf: 'center',
    paddingHorizontal: spacing.pageX,
    gap: spacing.md
  },
  hero: {
    marginBottom: spacing.sm
  },
  title: {
    color: colors.ink,
    fontFamily: 'SpaceGrotesk_700Bold',
    fontSize: 34,
    letterSpacing: -0.6,
    marginBottom: 10
  },
  sub: {
    color: colors.muted,
    fontSize: 16,
    lineHeight: 24,
    fontFamily: 'SpaceGrotesk_400Regular',
    marginBottom: 8
  },
  offlineHint: {
    color: colors.accent2,
    fontSize: 13,
    fontFamily: 'SpaceGrotesk_400Regular',
    opacity: 0.9
  },
  grid: {
    gap: spacing.cardGap
  },
  hint: {
    ...typography.caption,
    color: colors.muted,
    textAlign: 'center'
  },
  noteCard: {
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: spacing.cardRadius,
    padding: spacing.cardPadding,
    gap: 6
  },
  noteTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  noteCopy: {
    ...typography.caption,
    color: colors.muted,
    lineHeight: 20
  }
});
