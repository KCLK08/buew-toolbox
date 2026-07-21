import { StyleSheet, View } from 'react-native';

import { spacing } from '../../constants/theme';
import type { SiteReportColumn } from '../../native/sitereport/db/database';
import { PrimaryButton, TextField } from '../mobile';

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
        placeholder={column.name}
      />
    </View>
  );
}

const STATUS_OPTIONS = [
  { value: 'offen', label: 'Offen', emoji: '🟠', variant: 'secondary' as const },
  { value: 'bearbeitung', label: 'Bearbeitung', emoji: '🔵', variant: 'secondary' as const },
  { value: 'erledigt', label: 'Erledigt', emoji: '🟢', variant: 'secondary' as const }
];

type StatusProps = {
  value: string;
  onChange: (value: string) => void;
};

export function StatusFieldStep({ value, onChange }: StatusProps) {
  return (
    <View style={styles.statusWrap}>
      {STATUS_OPTIONS.map((opt) => (
        <PrimaryButton
          key={opt.value}
          label={`${opt.emoji} ${opt.label}`}
          variant={value === opt.value ? 'primary' : 'secondary'}
          onPress={() => onChange(opt.value)}
        />
      ))}
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: {
    gap: spacing.md
  },
  statusWrap: {
    gap: spacing.sm
  }
});
