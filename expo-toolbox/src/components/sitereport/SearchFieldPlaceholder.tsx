import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';

type Props = {
  placeholder?: string;
  onPress?: () => void;
};

export function SearchFieldPlaceholder({
  placeholder = 'Suchen…',
  onPress
}: Props) {
  return (
    <Pressable
      accessibilityRole="search"
      onPress={onPress}
      disabled={!onPress}
      style={({ pressed }) => [styles.field, pressed && onPress ? styles.pressed : null]}
    >
      <Text style={styles.icon}>🔍</Text>
      <Text style={styles.placeholder}>{placeholder}</Text>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  field: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: spacing.touchMin + 4,
    backgroundColor: colors.panelElevated,
    borderRadius: spacing.inputRadius,
    borderWidth: 1,
    borderColor: colors.border,
    paddingHorizontal: spacing.md,
    marginBottom: spacing.sm
  },
  pressed: {
    opacity: 0.9
  },
  icon: {
    fontSize: 16
  },
  placeholder: {
    ...typography.body,
    color: colors.muted,
    flex: 1
  }
});
