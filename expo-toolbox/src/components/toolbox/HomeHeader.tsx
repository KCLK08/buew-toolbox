import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import type { HomeLayoutTier } from './homeLayout';

type Props = {
  tier?: HomeLayoutTier;
};

export function HomeHeader({ tier = 'relaxed' }: Props) {
  const compact = tier !== 'relaxed';
  const dense = tier === 'dense';

  return (
    <View style={[styles.root, compact ? styles.rootCompact : null, dense ? styles.rootDense : null]}>
      <Text
        style={[styles.title, compact ? styles.titleCompact : null, dense ? styles.titleDense : null]}
        numberOfLines={dense ? 1 : undefined}
      >
        BÜW Toolbox
      </Text>
      <Text
        style={[
          styles.subtitle,
          compact ? styles.subtitleCompact : null,
          dense ? styles.subtitleDense : null
        ]}
        numberOfLines={dense ? 2 : undefined}
      >
        Digitale Werkzeuge für eine strukturierte Baustellendokumentation
      </Text>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.xs,
    paddingBottom: spacing.md
  },
  rootCompact: {
    paddingBottom: spacing.sm
  },
  rootDense: {
    gap: spacing.xxs,
    paddingBottom: spacing.xs
  },
  title: {
    ...typography.display,
    color: colors.ink,
    fontSize: 30,
    lineHeight: 36,
    letterSpacing: -0.5
  },
  titleCompact: {
    fontSize: 26,
    lineHeight: 30
  },
  titleDense: {
    fontSize: 24,
    lineHeight: 28
  },
  subtitle: {
    ...typography.body,
    color: colors.muted,
    fontSize: 15,
    lineHeight: 22,
    maxWidth: 520
  },
  subtitleCompact: {
    fontSize: 14,
    lineHeight: 20
  },
  subtitleDense: {
    fontSize: 13,
    lineHeight: 18
  }
});
