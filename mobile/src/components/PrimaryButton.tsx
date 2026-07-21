import { colors } from '@buew/shared';
import { ActivityIndicator, Pressable, StyleSheet, Text, type PressableProps } from 'react-native';

type PrimaryButtonProps = PressableProps & {
  label: string;
  loading?: boolean;
};

export function PrimaryButton({ label, loading, disabled, ...props }: PrimaryButtonProps) {
  const isDisabled = Boolean(disabled || loading);
  return (
    <Pressable
      {...props}
      disabled={isDisabled}
      style={[styles.button, isDisabled && styles.disabled]}
    >
      {loading ? (
        <ActivityIndicator color={colors.white} />
      ) : (
        <Text style={styles.label}>{label}</Text>
      )}
    </Pressable>
  );
}

const styles = StyleSheet.create({
  button: {
    backgroundColor: colors.accent,
    borderRadius: 12,
    paddingVertical: 14,
    alignItems: 'center',
    justifyContent: 'center'
  },
  disabled: {
    opacity: 0.6
  },
  label: {
    color: colors.white,
    fontFamily: 'SpaceGrotesk_700Bold',
    fontSize: 16
  }
});
