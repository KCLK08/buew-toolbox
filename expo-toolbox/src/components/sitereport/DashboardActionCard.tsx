import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, shadows, spacing, typography } from '../../constants/theme';

type Props = {
  title: string;
  subtitle?: string;
  icon?: string;
  accent?: boolean;
  onPress: () => void;
};

export function DashboardActionCard({ title, subtitle, icon, accent, onPress }: Props) {
  return (
    <Pressable
      onPress={onPress}
      style={({ pressed }) => [styles.card, accent ? styles.accent : null, pressed ? styles.pressed : null]}
    >
      {icon ? <Text style={styles.icon}>{icon}</Text> : null}
      <Text style={[styles.title, accent ? styles.titleAccent : null]}>{title}</Text>
      {subtitle ? (
        <Text style={[styles.subtitle, accent ? styles.subtitleAccent : null]}>{subtitle}</Text>
      ) : null}
    </Pressable>
  );
}

const styles = StyleSheet.create({
  card: {
    flex: 1,
    minWidth: '46%',
    backgroundColor: colors.panelElevated,
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    padding: spacing.cardPadding,
    gap: spacing.xs,
    ...shadows.card
  },
  accent: {
    backgroundColor: colors.accent,
    borderColor: colors.accent
  },
  pressed: {
    opacity: 0.92,
    transform: [{ scale: 0.98 }]
  },
  icon: {
    fontSize: 28
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  titleAccent: {
    color: colors.white
  },
  subtitle: {
    ...typography.caption,
    color: colors.muted
  },
  subtitleAccent: {
    color: 'rgba(255,255,255,0.85)'
  }
});
