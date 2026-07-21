import { Pressable, StyleSheet, Text } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, shadows, spacing, typography } from '../../constants/theme';

type Props = {
  label?: string;
  onPress: () => void;
  accessibilityLabel?: string;
};

export function Fab({ label = '+', onPress, accessibilityLabel }: Props) {
  const insets = useSafeAreaInsets();
  return (
    <Pressable
      accessibilityRole="button"
      accessibilityLabel={accessibilityLabel || 'Neue Aktion'}
      onPress={onPress}
      style={({ pressed }) => [
        styles.fab,
        shadows.fab,
        { bottom: Math.max(insets.bottom, 12) + 64 },
        pressed ? styles.pressed : null
      ]}
    >
      <Text style={styles.label}>{label}</Text>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  fab: {
    position: 'absolute',
    right: spacing.pageX,
    width: spacing.fabSize,
    height: spacing.fabSize,
    borderRadius: spacing.fabSize / 2,
    backgroundColor: colors.accent,
    alignItems: 'center',
    justifyContent: 'center'
  },
  pressed: {
    backgroundColor: colors.accentPressed,
    transform: [{ scale: 0.96 }]
  },
  label: {
    ...typography.title,
    color: colors.white,
    marginTop: -2
  }
});
