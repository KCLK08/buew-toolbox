import { useEffect, useRef } from 'react';
import { Animated, Image, StyleSheet, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../constants/theme';
import { systemBottomInset } from '../navigation/systemInsets';

const appIcon = require('../../assets/icon.png');

type Props = {
  onFinish: () => void;
  minDurationMs?: number;
};

export function AppLoadingScreen({ onFinish, minDurationMs = 1600 }: Props) {
  const insets = useSafeAreaInsets();
  const fade = useRef(new Animated.Value(0)).current;
  const scale = useRef(new Animated.Value(0.92)).current;
  const pulse = useRef(new Animated.Value(0.4)).current;

  useEffect(() => {
    Animated.parallel([
      Animated.timing(fade, { toValue: 1, duration: 520, useNativeDriver: true }),
      Animated.spring(scale, { toValue: 1, friction: 7, tension: 90, useNativeDriver: true })
    ]).start();

    const pulseLoop = Animated.loop(
      Animated.sequence([
        Animated.timing(pulse, { toValue: 1, duration: 700, useNativeDriver: true }),
        Animated.timing(pulse, { toValue: 0.35, duration: 700, useNativeDriver: true })
      ])
    );
    pulseLoop.start();

    const timer = setTimeout(() => {
      Animated.timing(fade, { toValue: 0, duration: 320, useNativeDriver: true }).start(({ finished }) => {
        if (finished) onFinish();
      });
    }, minDurationMs);

    return () => {
      clearTimeout(timer);
      pulseLoop.stop();
    };
  }, [fade, minDurationMs, onFinish, pulse, scale]);

  return (
    <View style={[styles.root, { paddingTop: insets.top, paddingBottom: systemBottomInset(insets) }]}>
      <Animated.View style={[styles.content, { opacity: fade, transform: [{ scale }] }]}>
        <View style={styles.logoWrap}>
          <Image
            source={appIcon}
            style={styles.logo}
            resizeMode="contain"
            accessibilityLabel="BÜW Toolbox"
          />
        </View>
        <Text style={styles.title}>BÜW Toolbox</Text>
        <Text style={styles.subtitle}>Digitale Baustellendokumentation</Text>
        <View style={styles.loaderTrack}>
          <Animated.View style={[styles.loaderBar, { opacity: pulse }]} />
        </View>
      </Animated.View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg,
    alignItems: 'center',
    justifyContent: 'center',
    paddingHorizontal: spacing.xl
  },
  content: {
    alignItems: 'center',
    gap: spacing.sm,
    width: '100%',
    maxWidth: 320
  },
  logoWrap: {
    width: 84,
    height: 84,
    borderRadius: 22,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border,
    alignItems: 'center',
    justifyContent: 'center',
    marginBottom: spacing.xs,
    padding: spacing.sm,
    overflow: 'hidden',
    shadowColor: '#1A1916',
    shadowOpacity: 0.1,
    shadowRadius: 16,
    shadowOffset: { width: 0, height: 6 },
    elevation: 3
  },
  logo: {
    width: 56,
    height: 56
  },
  title: {
    ...typography.display,
    color: colors.ink,
    fontSize: 30,
    textAlign: 'center'
  },
  subtitle: {
    ...typography.body,
    color: colors.muted,
    textAlign: 'center',
    fontSize: 15,
    lineHeight: 22
  },
  loaderTrack: {
    marginTop: spacing.lg,
    width: 120,
    height: 4,
    borderRadius: 999,
    backgroundColor: colors.border,
    overflow: 'hidden'
  },
  loaderBar: {
    flex: 1,
    borderRadius: 999,
    backgroundColor: colors.accent
  }
});
