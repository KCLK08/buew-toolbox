import { ActivityIndicator, Pressable, StyleSheet, Text, type ViewStyle } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';

type Variant = 'primary' | 'secondary' | 'ghost' | 'danger';

type Props = {
  label: string;
  onPress: () => void;
  variant?: Variant;
  disabled?: boolean;
  loading?: boolean;
  style?: ViewStyle;
};

export function PrimaryButton({
  label,
  onPress,
  variant = 'primary',
  disabled,
  loading,
  style
}: Props) {
  const isDisabled = disabled || loading;
  return (
    <Pressable
      accessibilityRole="button"
      disabled={isDisabled}
      onPress={onPress}
      style={({ pressed }) => [
        styles.base,
        styles[variant],
        pressed && !isDisabled ? styles.pressed : null,
        isDisabled ? styles.disabled : null,
        style
      ]}
    >
      {loading ? (
        <ActivityIndicator color={variant === 'primary' || variant === 'danger' ? colors.white : colors.accent} />
      ) : (
        <Text
          style={[
            styles.label,
            variant === 'primary' || variant === 'danger' ? styles.labelOnAccent : styles.labelMuted
          ]}
        >
          {label}
        </Text>
      )}
    </Pressable>
  );
}

const styles = StyleSheet.create({
  base: {
    minHeight: spacing.touchMin + 8,
    borderRadius: spacing.buttonRadius,
    alignItems: 'center',
    justifyContent: 'center',
    paddingHorizontal: spacing.lg,
    paddingVertical: spacing.sm
  },
  primary: {
    backgroundColor: colors.accent
  },
  secondary: {
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.borderStrong
  },
  ghost: {
    backgroundColor: 'transparent'
  },
  danger: {
    backgroundColor: colors.danger
  },
  pressed: {
    opacity: 0.9,
    transform: [{ scale: 0.98 }]
  },
  disabled: {
    opacity: 0.5
  },
  label: {
    ...typography.button
  },
  labelOnAccent: {
    color: colors.white
  },
  labelMuted: {
    color: colors.ink
  }
});
