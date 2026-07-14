import { Pressable, StyleSheet, Switch, Text, TextInput, View } from 'react-native';

import { colors } from '@/theme/colors';
import type { SetupField } from '@/types';
import { normalizeClockTime } from '@/lib/time-format';

interface FieldInputProps {
  field: SetupField;
  value: string | boolean | undefined;
  onChange: (value: string | boolean) => void;
  compact?: boolean;
}

export function FieldInput({ field, value, onChange, compact = false }: FieldInputProps) {
  const label = field.label || field.fieldName;

  if (field.type === 'checkbox') {
    return (
      <Pressable style={[styles.row, compact && styles.compact]} onPress={() => onChange(!(value === true))}>
        <Switch value={value === true} onValueChange={onChange} trackColor={{ true: colors.primary }} />
        <Text style={styles.label}>{label}</Text>
      </Pressable>
    );
  }

  if (field.type === 'dropdown' && field.options.length > 0) {
    return (
      <View style={[styles.block, compact && styles.compact]}>
        <Text style={styles.label}>{label}</Text>
        <View style={styles.optionList}>
          {field.options.map((option) => {
            const selected = String(value ?? '') === option;
            return (
              <Pressable
                key={option}
                style={[styles.option, selected && styles.optionSelected]}
                onPress={() => onChange(option)}
              >
                <Text style={[styles.optionText, selected && styles.optionTextSelected]}>{option}</Text>
              </Pressable>
            );
          })}
        </View>
      </View>
    );
  }

  const isMultiline = ['Text63', 'Text64', 'Text66', 'Text67', 'Text70'].includes(field.fieldName);

  return (
    <View style={[styles.block, compact && styles.compact]}>
      <Text style={styles.label}>{label}</Text>
      <TextInput
        style={[styles.input, isMultiline && styles.multiline]}
        value={String(value ?? '')}
        onChangeText={(text) => {
          const isTimeField = field.fieldName && ['Text21', 'Text22', 'Text23', 'Text24'].some((n) => field.fieldName.includes(n));
          onChange(isTimeField ? normalizeClockTime(text) : text);
        }}
        multiline={isMultiline}
        placeholder={label}
        placeholderTextColor={colors.textMuted}
      />
    </View>
  );
}

const styles = StyleSheet.create({
  block: {
    marginBottom: 14,
  },
  compact: {
    marginBottom: 8,
  },
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 12,
    marginBottom: 14,
    paddingVertical: 4,
  },
  label: {
    color: colors.text,
    fontSize: 14,
    fontWeight: '600',
    marginBottom: 6,
  },
  input: {
    backgroundColor: colors.surface,
    borderColor: colors.border,
    borderRadius: 10,
    borderWidth: 1,
    color: colors.text,
    fontSize: 16,
    minHeight: 44,
    paddingHorizontal: 12,
    paddingVertical: 10,
  },
  multiline: {
    minHeight: 100,
    textAlignVertical: 'top',
  },
  optionList: {
    gap: 8,
  },
  option: {
    backgroundColor: colors.surface,
    borderColor: colors.border,
    borderRadius: 10,
    borderWidth: 1,
    paddingHorizontal: 12,
    paddingVertical: 10,
  },
  optionSelected: {
    backgroundColor: '#e8f4f2',
    borderColor: colors.primary,
  },
  optionText: {
    color: colors.text,
    fontSize: 15,
  },
  optionTextSelected: {
    color: colors.primary,
    fontWeight: '600',
  },
});
