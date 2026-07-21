import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, shadows, spacing, typography } from '../../constants/theme';
import { Card } from '../mobile/Card';
import { SelectionCheckbox } from './SelectionCheckbox';

type Props = {
  title: string;
  subtitle?: string;
  date?: string;
  entryCount?: number;
  photoCount?: number;
  meta?: string;
  onPress?: () => void;
  trailing?: React.ReactNode;
  selected?: boolean;
  onSelectToggle?: () => void;
};

export function ProtocolCard({
  title,
  subtitle,
  date,
  entryCount,
  photoCount,
  meta,
  onPress,
  trailing,
  selected,
  onSelectToggle
}: Props) {
  const chips = [
    date ? { label: date, icon: '📅' } : null,
    entryCount !== undefined ? { label: `${entryCount} Einträge`, icon: '✏️' } : null,
    photoCount !== undefined ? { label: `${photoCount} Fotos`, icon: '📷' } : null
  ].filter(Boolean) as { label: string; icon: string }[];

  return (
    <Pressable
      accessibilityRole="button"
      onPress={onPress}
      disabled={!onPress}
      style={({ pressed }) => [pressed && onPress ? styles.pressed : null]}
    >
      <Card style={selected ? { ...styles.card, ...styles.selected } : styles.card}>
        <View style={styles.row}>
          {onSelectToggle ? <SelectionCheckbox selected={selected} onToggle={onSelectToggle} /> : null}
          <View style={styles.body}>
            <Text style={styles.title} numberOfLines={2}>
              {title}
            </Text>
            {subtitle ? (
              <Text style={styles.subtitle} numberOfLines={1}>
                {subtitle}
              </Text>
            ) : null}
            {chips.length > 0 ? (
              <View style={styles.chips}>
                {chips.map((chip) => (
                  <View key={chip.label} style={styles.chip}>
                    <Text style={styles.chipIcon}>{chip.icon}</Text>
                    <Text style={styles.chipLabel}>{chip.label}</Text>
                  </View>
                ))}
              </View>
            ) : meta ? (
              <Text style={styles.meta}>{meta}</Text>
            ) : null}
          </View>
          {trailing}
        </View>
      </Card>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  card: {
    marginBottom: spacing.sm,
    ...shadows.card
  },
  selected: {
    borderColor: colors.accent,
    borderWidth: 2
  },
  pressed: {
    opacity: 0.94,
    transform: [{ scale: 0.995 }]
  },
  row: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.sm
  },
  body: {
    flex: 1,
    gap: 4
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink,
    fontSize: 17
  },
  subtitle: {
    ...typography.body,
    color: colors.muted
  },
  chips: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs,
    marginTop: spacing.xs
  },
  chip: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 4,
    backgroundColor: colors.bg,
    borderRadius: 999,
    paddingHorizontal: spacing.sm,
    paddingVertical: 4
  },
  chipIcon: {
    fontSize: 12
  },
  chipLabel: {
    ...typography.caption,
    color: colors.accent2
  },
  meta: {
    ...typography.caption,
    color: colors.accent2,
    marginTop: 2
  }
});
