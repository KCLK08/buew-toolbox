import { colors } from '@buew/shared';
import { useState } from 'react';
import {
  Pressable,
  StyleSheet,
  Text,
  TextInput,
  View,
  type TextInputProps
} from 'react-native';

type FieldProps = TextInputProps & {
  label: string;
  error?: string;
};

export function TextField({ label, error, style, ...props }: FieldProps) {
  return (
    <View style={styles.field}>
      <Text style={styles.label}>{label}</Text>
      <TextInput
        {...props}
        placeholderTextColor={colors.muted}
        style={[styles.input, style]}
      />
      {error ? <Text style={styles.error}>{error}</Text> : null}
    </View>
  );
}

export function PasswordField({ label, error, style, ...props }: FieldProps) {
  const [visible, setVisible] = useState(false);

  return (
    <View style={styles.field}>
      <Text style={styles.label}>{label}</Text>
      <View style={styles.passwordRow}>
        <TextInput
          {...props}
          secureTextEntry={!visible}
          placeholderTextColor={colors.muted}
          style={[styles.input, styles.passwordInput, style]}
        />
        <Pressable
          accessibilityRole="button"
          onPress={() => setVisible((v) => !v)}
          style={styles.toggle}
        >
          <Text style={styles.toggleText}>{visible ? 'Verbergen' : 'Anzeigen'}</Text>
        </Pressable>
      </View>
      {error ? <Text style={styles.error}>{error}</Text> : null}
    </View>
  );
}

const styles = StyleSheet.create({
  field: {
    gap: 6
  },
  label: {
    fontFamily: 'SpaceGrotesk_600SemiBold',
    color: colors.ink,
    fontSize: 14
  },
  input: {
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.white,
    borderRadius: 12,
    paddingHorizontal: 14,
    paddingVertical: 12,
    fontFamily: 'SpaceGrotesk_400Regular',
    color: colors.ink,
    fontSize: 16
  },
  passwordRow: {
    flexDirection: 'row',
    gap: 8,
    alignItems: 'center'
  },
  passwordInput: {
    flex: 1
  },
  toggle: {
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.white,
    borderRadius: 12,
    paddingHorizontal: 10,
    paddingVertical: 12
  },
  toggleText: {
    fontFamily: 'SpaceGrotesk_600SemiBold',
    color: colors.accent2,
    fontSize: 13
  },
  error: {
    color: colors.danger,
    fontFamily: 'SpaceGrotesk_400Regular',
    fontSize: 13
  }
});
