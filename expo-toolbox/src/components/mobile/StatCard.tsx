import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, shadows, spacing, typography } from '../../constants/theme';
import { Card } from './Card';

type Props = {
  title: string;
  value: string;
  icon?: string;
  onPress?: () => void;
  tone?: 'default' | 'accent' | 'warning';
};

export function StatCard({ title, value, icon, onPress, tone = 'default' }: Props) {
  const content = (
    <Card
      style={StyleSheet.flatten([
        styles.card,
        tone === 'accent' ? styles.accent : null,
        tone === 'warning' ? styles.warning : null
      ])}
    >
      <View style={styles.top}>
        {icon ? <Text style={styles.icon}>{icon}</Text> : null}
        <Text style={styles.title}>{title}</Text>
      </View>
      <Text style={styles.value}>{value}</Text>
    </Card>
  );

  if (!onPress) return content;

  return (
    <Pressable accessibilityRole="button" onPress={onPress} style={({ pressed }) => [styles.pressable, pressed ? styles.pressed : null]}>
      {content}
    </Pressable>
  );
}

const styles = StyleSheet.create({
  pressable: {
    flex: 1,
    minWidth: '46%'
  },
  pressed: {
    opacity: 0.92,
    transform: [{ scale: 0.98 }]
  },
  card: {
    ...shadows.card,
    minHeight: 108,
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  accent: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  warning: {
    borderColor: colors.warning
  },
  top: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs
  },
  icon: {
    fontSize: 18
  },
  title: {
    ...typography.label,
    color: colors.muted,
    flexShrink: 1
  },
  value: {
    ...typography.display,
    color: colors.ink,
    fontSize: 26,
    lineHeight: 32
  }
});
