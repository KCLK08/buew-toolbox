import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import type { SiteReportColumn } from '../../native/sitereport/db/database';
import { TextField } from '../mobile';
import { hapticSelection } from '../../lib/haptics';

type Props = {
  column: SiteReportColumn;
  value: string;
  onChange: (value: string) => void;
};

export function FieldStep({ column, value, onChange }: Props) {
  return (
    <View style={styles.wrap}>
      <TextField
        label={column.name}
        value={value}
        onChangeText={onChange}
        keyboardType={column.type === 'number' ? 'decimal-pad' : 'default'}
        multiline={column.type === 'text'}
        placeholder={`${column.name} eingeben…`}
        autoFocus
      />
    </View>
  );
}

const STATUS_OPTIONS = [
  { value: 'offen', label: 'Offen', emoji: '🟠' },
  { value: 'bearbeitung', label: 'Bearbeitung', emoji: '🔵' },
  { value: 'erledigt', label: 'Erledigt', emoji: '🟢' }
];

type StatusProps = {
  value: string;
  onChange: (value: string) => void;
};

export function StatusFieldStep({ value, onChange }: StatusProps) {
  return (
    <View style={styles.statusWrap}>
      {STATUS_OPTIONS.map((opt) => {
        const active = value === opt.value;
        return (
          <Pressable
            key={opt.value}
            accessibilityRole="button"
            onPress={() => {
              void hapticSelection();
              onChange(opt.value);
            }}
            style={[styles.statusOption, active ? styles.statusActive : null]}
          >
            <Text style={[styles.statusLabel, active ? styles.statusLabelActive : null]}>
              {opt.emoji} {opt.label}
            </Text>
          </Pressable>
        );
      })}
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: {
    gap: spacing.md,
    minHeight: 120
  },
  statusWrap: {
    gap: spacing.sm
  },
  statusOption: {
    minHeight: spacing.touchMin + 8,
    borderRadius: spacing.buttonRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    alignItems: 'center',
    justifyContent: 'center'
  },
  statusActive: {
    backgroundColor: colors.accent,
    borderColor: colors.accent
  },
  statusLabel: {
    ...typography.button,
    color: colors.ink
  },
  statusLabelActive: {
    color: colors.white
  }
});
