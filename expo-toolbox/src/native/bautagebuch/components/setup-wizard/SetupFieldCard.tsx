import { Pressable, StyleSheet, Switch, Text, View } from 'react-native';

import { TextField } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import {
  checkboxBehaviorHint,
  isCheckboxField,
  readCheckboxDefault,
  writeCheckboxDefault
} from '../../lib/setup-field-hints.js';
import type { DetectedField, SetupFieldConfig } from '../../types';

type Props = {
  field: SetupFieldConfig;
  detectedFields?: DetectedField[];
  expanded?: boolean;
  readOnly?: boolean;
  onPress?: () => void;
  onChange: (patch: Partial<SetupFieldConfig>) => void;
};

function StatusPills({ items }: { items: string[] }) {
  if (items.length === 0) {
    return <Text style={styles.pillNeutral}>Standard</Text>;
  }
  return (
    <View style={styles.pillRow}>
      {items.map((item) => (
        <View key={item} style={styles.pill}>
          <Text style={styles.pillText}>{item}</Text>
        </View>
      ))}
    </View>
  );
}

function SettingRow({
  title,
  hint,
  value,
  onValueChange,
  disabled = false
}: {
  title: string;
  hint: string;
  value: boolean;
  onValueChange: (next: boolean) => void;
  disabled?: boolean;
}) {
  return (
    <View style={styles.settingRow}>
      <View style={styles.settingCopy}>
        <Text style={styles.settingTitle}>{title}</Text>
        <Text style={styles.settingHint}>{hint}</Text>
      </View>
      <Switch
        value={value}
        onValueChange={onValueChange}
        disabled={disabled}
        trackColor={{ false: colors.border, true: colors.accent }}
        thumbColor={colors.panel}
      />
    </View>
  );
}

export function SetupFieldCard({
  field,
  detectedFields = [],
  expanded = false,
  readOnly = false,
  onPress,
  onChange
}: Props) {
  const title = field.label || field.fieldName || field.fieldId;
  const checkboxField = isCheckboxField(field, detectedFields);

  return (
    <Pressable
      style={[styles.card, expanded ? styles.cardActive : null]}
      onPress={onPress}
      disabled={readOnly && !onPress}
    >
      <View style={styles.header}>
        <View style={styles.heading}>
          <Text style={styles.title} numberOfLines={2}>
            {title}
          </Text>
          <Text style={styles.meta}>
            {field.fieldName || field.fieldId}
            {field.page ? ` · Seite ${field.page}` : ''}
          </Text>
        </View>
        {onPress ? <Text style={styles.chevron}>{expanded ? '▲' : '▼'}</Text> : null}
      </View>

      {!expanded ? (
        <StatusPills
          items={[
            field.required ? 'Pflicht' : null,
            field.skipped ? 'Ausgeblendet' : null,
            field.multiline ? 'Mehrzeilig' : null,
            field.defaultValue ? 'Standardwert' : null
          ].filter((item): item is string => Boolean(item))}
        />
      ) : (
        <View style={styles.body}>
          <TextField
            label="Anzeigename im Assistenten"
            value={field.label || ''}
            editable={!readOnly}
            onChangeText={(label) => onChange({ label })}
          />

          {checkboxField ? (
            <SettingRow
              title="Standard: aktiviert (Ja)"
              hint={checkboxBehaviorHint(field.fieldName)}
              value={readCheckboxDefault(field)}
              disabled={readOnly}
              onValueChange={(value) => onChange({ defaultValue: writeCheckboxDefault(value) })}
            />
          ) : (
            <TextField
              label="Standardtext zum Vorausfüllen"
              value={field.defaultValue || ''}
              editable={!readOnly}
              onChangeText={(defaultValue) => onChange({ defaultValue })}
              placeholder="Optional — wird beim Start des BTB gesetzt"
            />
          )}

          <SettingRow
            title="Pflichtfeld"
            hint="Muss vor dem Export ausgefüllt sein"
            value={Boolean(field.required)}
            disabled={readOnly}
            onValueChange={(required) => onChange({ required })}
          />
          <SettingRow
            title="Im Assistenten ausblenden"
            hint="Feld wird nicht angezeigt, bleibt aber im PDF"
            value={Boolean(field.skipped)}
            disabled={readOnly}
            onValueChange={(skipped) => onChange({ skipped })}
          />
          {!checkboxField ? (
            <SettingRow
              title="Mehrzeiliges Eingabefeld"
              hint="Größeres Textfeld für längere Einträge"
              value={Boolean(field.multiline)}
              disabled={readOnly}
              onValueChange={(multiline) => onChange({ multiline })}
            />
          ) : null}

          <TextField
            label="Hilfetext"
            value={field.hint || ''}
            editable={!readOnly}
            onChangeText={(hint) => onChange({ hint })}
            multiline
          />
        </View>
      )}
    </Pressable>
  );
}

const styles = StyleSheet.create({
  card: {
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel,
    padding: spacing.sm,
    gap: spacing.sm
  },
  cardActive: {
    borderColor: colors.accent,
    backgroundColor: 'rgba(47, 111, 237, 0.06)'
  },
  header: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  heading: {
    flex: 1,
    gap: 4
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  meta: {
    ...typography.caption,
    color: colors.muted
  },
  chevron: {
    ...typography.caption,
    color: colors.muted,
    marginTop: 2
  },
  body: {
    gap: spacing.xs,
    paddingTop: spacing.xs,
    borderTopWidth: 1,
    borderTopColor: colors.border
  },
  settingRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm,
    paddingVertical: spacing.xs
  },
  settingCopy: {
    flex: 1,
    gap: 2
  },
  settingTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  settingHint: {
    ...typography.caption,
    color: colors.muted
  },
  pillRow: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xxs
  },
  pill: {
    borderRadius: 999,
    paddingHorizontal: spacing.sm,
    paddingVertical: 4,
    backgroundColor: colors.badgeBg
  },
  pillText: {
    ...typography.caption,
    color: colors.accent
  },
  pillNeutral: {
    ...typography.caption,
    color: colors.muted
  }
});
