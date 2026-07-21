import { ReactNode, useEffect, useRef } from 'react';
import { Animated, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';

type Props = {
  step: number;
  total: number;
  title: string;
  children: ReactNode;
};

export function WizardStep({ step, total, title, children }: Props) {
  const fade = useRef(new Animated.Value(1)).current;
  const slide = useRef(new Animated.Value(0)).current;

  useEffect(() => {
    fade.setValue(0);
    slide.setValue(12);
    Animated.parallel([
      Animated.timing(fade, { toValue: 1, duration: 240, useNativeDriver: true }),
      Animated.spring(slide, { toValue: 0, useNativeDriver: true, speed: 18, bounciness: 4 })
    ]).start();
  }, [fade, slide, step, title]);

  return (
    <View style={styles.wrap}>
      <Text style={styles.progress}>
        Schritt {step} von {total}
      </Text>
      <View style={styles.dots}>
        {Array.from({ length: total }, (_, index) => (
          <View
            key={index}
            style={[styles.dot, index < step ? styles.dotDone : null, index === step - 1 ? styles.dotActive : null]}
          />
        ))}
      </View>
      <Text style={styles.title}>{title}</Text>
      <View style={styles.barTrack}>
        <View style={[styles.barFill, { width: `${(step / total) * 100}%` }]} />
      </View>
      <Animated.View style={{ opacity: fade, transform: [{ translateY: slide }] }}>{children}</Animated.View>
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: {
    gap: spacing.md
  },
  progress: {
    ...typography.label,
    color: colors.accent,
    textTransform: 'uppercase',
    letterSpacing: 0.6
  },
  dots: {
    flexDirection: 'row',
    gap: spacing.xs
  },
  dot: {
    flex: 1,
    height: 4,
    borderRadius: 999,
    backgroundColor: colors.border
  },
  dotDone: {
    backgroundColor: colors.accent
  },
  dotActive: {
    backgroundColor: colors.accent,
    transform: [{ scaleY: 1.4 }]
  },
  title: {
    ...typography.title,
    color: colors.ink
  },
  barTrack: {
    height: 4,
    borderRadius: 999,
    backgroundColor: colors.border,
    overflow: 'hidden'
  },
  barFill: {
    height: '100%',
    backgroundColor: colors.accent,
    borderRadius: 999
  }
});
