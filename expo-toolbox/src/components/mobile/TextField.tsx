import { StyleSheet, Text, TextInput, View, type TextInputProps } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';

type Props = TextInputProps & {
  label: string;
  hint?: string;
};

export function TextField({ label, hint, style, ...rest }: Props) {
  return (
    <View style={styles.wrap}>
      <Text style={styles.label}>{label}</Text>
      <TextInput
        placeholderTextColor={colors.muted}
        style={[styles.input, style]}
        {...rest}
      />
      {hint ? <Text style={styles.hint}>{hint}</Text> : null}
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: {
    gap: spacing.xs,
    marginBottom: spacing.sm
  },
  label: {
    ...typography.label,
    color: colors.ink
  },
  input: {
    minHeight: spacing.touchMin + 8,
    borderWidth: 1,
    borderColor: colors.borderStrong,
    borderRadius: spacing.inputRadius,
    backgroundColor: colors.panelElevated,
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
    color: colors.ink,
    ...typography.body
  },
  hint: {
    ...typography.caption,
    color: colors.muted
  }
});
