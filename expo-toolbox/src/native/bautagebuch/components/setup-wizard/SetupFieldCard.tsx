import { StyleSheet, Switch, Text, View } from 'react-native';

import { TextField } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import type { SetupFieldConfig } from '../../types';

type Props = {
  field: SetupFieldConfig;
  readOnly?: boolean;
  onChange: (patch: Partial<SetupFieldConfig>) => void;
};

export function SetupFieldCard({ field, readOnly = false, onChange }: Props) {
  const title = field.label || field.fieldName || field.fieldId;

  return (
    <View style={[styles.card, shadows.card]}>
      <Text style={styles.title}>{title}</Text>
      <Text style={styles.meta}>
        {field.fieldName || field.fieldId}
        {field.page ? ` · Seite ${field.page}` : ''}
      </Text>

      <TextField
        label="Label"
        value={field.label || ''}
        editable={!readOnly}
        onChangeText={(label) => onChange({ label })}
      />

      <View style={styles.toggleRow}>
        <Text style={styles.toggleLabel}>Pflichtfeld</Text>
        <Switch
          value={Boolean(field.required)}
          disabled={readOnly}
          onValueChange={(required) => onChange({ required })}
          trackColor={{ false: colors.border, true: colors.accent }}
        />
      </View>

      <View style={styles.toggleRow}>
        <Text style={styles.toggleLabel}>Überspringen</Text>
        <Switch
          value={Boolean(field.skipped)}
          disabled={readOnly}
          onValueChange={(skipped) => onChange({ skipped })}
          trackColor={{ false: colors.border, true: colors.accent }}
        />
      </View>

      <View style={styles.toggleRow}>
        <Text style={styles.toggleLabel}>Mehrzeilig</Text>
        <Switch
          value={Boolean(field.multiline)}
          disabled={readOnly}
          onValueChange={(multiline) => onChange({ multiline })}
          trackColor={{ false: colors.border, true: colors.accent }}
        />
      </View>

      <TextField
        label="Standardwert"
        value={field.defaultValue || ''}
        editable={!readOnly}
        onChangeText={(defaultValue) => onChange({ defaultValue })}
      />

      <TextField
        label="Hilfetext"
        value={field.hint || ''}
        editable={!readOnly}
        onChangeText={(hint) => onChange({ hint })}
        multiline
      />
    </View>
  );
}

const styles = StyleSheet.create({
  card: {
    backgroundColor: colors.panelElevated,
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    padding: spacing.cardPadding,
    gap: spacing.sm
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  meta: {
    ...typography.caption,
    color: colors.muted
  },
  toggleRow: {
    minHeight: spacing.touchMin,
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  toggleLabel: {
    ...typography.body,
    color: colors.ink
  }
});
