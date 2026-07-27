import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';

type Props = {
  compact?: boolean;
};

export function HomeHeader({ compact = false }: Props) {
  return (
    <View style={[styles.root, compact ? styles.rootCompact : null]}>
      <Text style={[styles.title, compact ? styles.titleCompact : null]}>BÜW Toolbox</Text>
      <Text style={[styles.subtitle, compact ? styles.subtitleCompact : null]}>
        Digitale Werkzeuge für eine strukturierte Baustellendokumentation
      </Text>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.xs,
    paddingBottom: spacing.sm
  },
  rootCompact: {
    paddingBottom: spacing.xs
  },
  title: {
    ...typography.display,
    color: colors.ink,
    fontSize: 32,
    letterSpacing: -0.5
  },
  titleCompact: {
    fontSize: 28,
    lineHeight: 32
  },
  subtitle: {
    ...typography.body,
    color: colors.muted,
    fontSize: 16,
    lineHeight: 24,
    maxWidth: 520
  },
  subtitleCompact: {
    fontSize: 14,
    lineHeight: 20
  }
});
