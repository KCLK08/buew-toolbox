import { ReactNode } from 'react';
import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import { Card } from './Card';

type Props = {
  title: string;
  subtitle?: string;
  meta?: string;
  trailing?: ReactNode;
  onPress?: () => void;
  badge?: ReactNode;
};

export function ListItem({ title, subtitle, meta, trailing, onPress, badge }: Props) {
  const content = (
    <Card style={styles.card} padded>
      <View style={styles.row}>
        <View style={styles.copy}>
          <Text style={styles.title} numberOfLines={2}>
            {title}
          </Text>
          {subtitle ? (
            <Text style={styles.subtitle} numberOfLines={2}>
              {subtitle}
            </Text>
          ) : null}
          {meta ? (
            <Text style={styles.meta} numberOfLines={1}>
              {meta}
            </Text>
          ) : null}
          {badge}
        </View>
        <View style={styles.trailing}>
          {trailing}
          {onPress ? <Text style={styles.chevron}>›</Text> : null}
        </View>
      </View>
    </Card>
  );

  if (!onPress) return content;

  return (
    <Pressable
      accessibilityRole="button"
      onPress={onPress}
      style={({ pressed }) => [pressed ? styles.pressed : null]}
    >
      {content}
    </Pressable>
  );
}

const styles = StyleSheet.create({
  card: {
    marginBottom: spacing.sm
  },
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm
  },
  copy: {
    flex: 1,
    gap: 4
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  subtitle: {
    ...typography.caption,
    color: colors.muted
  },
  meta: {
    ...typography.caption,
    color: colors.accent2
  },
  trailing: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs
  },
  chevron: {
    fontSize: 28,
    lineHeight: 28,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_400Regular'
  },
  pressed: {
    opacity: 0.92,
    transform: [{ scale: 0.995 }]
  }
});
