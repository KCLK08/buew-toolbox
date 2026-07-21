import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';

type Tone = 'neutral' | 'success' | 'warning' | 'danger' | 'info' | 'accent';

type Props = {
  label: string;
  tone?: Tone;
};

const toneStyles: Record<Tone, { bg: string; fg: string }> = {
  neutral: { bg: colors.border, fg: colors.ink },
  success: { bg: 'rgba(47, 107, 69, 0.14)', fg: colors.success },
  warning: { bg: 'rgba(154, 107, 18, 0.16)', fg: colors.warning },
  danger: { bg: 'rgba(161, 44, 36, 0.14)', fg: colors.danger },
  info: { bg: 'rgba(42, 95, 143, 0.14)', fg: colors.info },
  accent: { bg: colors.badgeBg, fg: colors.accent }
};

export function StatusBadge({ label, tone = 'neutral' }: Props) {
  const palette = toneStyles[tone];
  return (
    <View style={[styles.badge, { backgroundColor: palette.bg }]}>
      <Text style={[styles.label, { color: palette.fg }]}>{label}</Text>
    </View>
  );
}

const styles = StyleSheet.create({
  badge: {
    alignSelf: 'flex-start',
    borderRadius: 999,
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.xxs + 1
  },
  label: {
    ...typography.label,
    textTransform: 'capitalize'
  }
});
