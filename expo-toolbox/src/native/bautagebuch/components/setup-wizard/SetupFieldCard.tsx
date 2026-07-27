import { Pressable, StyleSheet, Switch, Text, View } from 'react-native';

import { TextField, SingleLineText } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import {
  checkboxBehaviorHint,
  isCheckboxField,
  readCheckboxDefault,
  resolveSetupFieldType,
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

function fieldTypeLabel(field: SetupFieldConfig, detectedFields: DetectedField[]): string {
  const type = resolveSetupFieldType(field, detectedFields);
  if (type === 'checkbox') return 'Checkbox';
  if (type === 'dropdown') return 'Dropdown';
  if (type === 'radio') return 'Radio';
  return 'Textfeld';
}

function ToggleRow({
  title,
  value,
  onValueChange,
  disabled = false
}: {
  title: string;
  value: boolean;
  onValueChange: (next: boolean) => void;
  disabled?: boolean;
}) {
  return (
    <View style={styles.toggleRow}>
      <Text style={styles.toggleTitle}>{title}</Text>
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
  const typeLabel = fieldTypeLabel(field, detectedFields);

  return (
    <Pressable
      style={[styles.card, expanded ? styles.cardActive : null]}
      onPress={onPress}
      disabled={readOnly && !onPress}
    >
      <View style={styles.header}>
        <View style={styles.heading}>
          <SingleLineText style={styles.title}>{title}</SingleLineText>
          <View style={styles.metaRow}>
            <Text style={styles.typeBadge}>{typeLabel}</Text>
            {field.page ? <Text style={styles.meta}>Seite {field.page}</Text> : null}
          </View>
        </View>
        {onPress ? <Text style={styles.chevron}>{expanded ? '▲' : '▼'}</Text> : null}
      </View>

      {!expanded ? (
        <View style={styles.collapsedFlags}>
          {field.required ? <Text style={styles.flagOn}>Pflichtfeld ✓</Text> : null}
          {field.skipped ? <Text style={styles.flagMuted}>Ausgeblendet</Text> : null}
          {field.multiline ? <Text style={styles.flagOn}>Mehrzeilig ✓</Text> : null}
          {!field.required && !field.skipped && !field.multiline ? (
            <Text style={styles.flagMuted}>Standard</Text>
          ) : null}
        </View>
      ) : (
        <View style={styles.body}>
          <TextField
            label="Anzeigename"
            value={field.label || ''}
            editable={!readOnly}
            onChangeText={(label) => onChange({ label })}
          />

          {checkboxField ? (
            <ToggleRow
              title="Standard: aktiviert (Ja)"
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
              placeholder="Optional"
            />
          )}

          <ToggleRow
            title="Pflichtfeld"
            value={Boolean(field.required)}
            disabled={readOnly}
            onValueChange={(required) => onChange({ required })}
          />
          <ToggleRow
            title="Im Assistenten ausblenden"
            value={Boolean(field.skipped)}
            disabled={readOnly}
            onValueChange={(skipped) => onChange({ skipped })}
          />
          {!checkboxField ? (
            <ToggleRow
              title="Mehrzeilig"
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

          <View style={styles.techBlock}>
            <Text style={styles.techHeading}>Technische Informationen</Text>
            <Text style={styles.techLine}>fieldName: {field.fieldName || '—'}</Text>
            <Text style={styles.techLine}>fieldId: {field.fieldId}</Text>
            <Text style={styles.techLine}>page: {field.page ?? '—'}</Text>
            <Text style={styles.techLine}>type: {field.type || typeLabel.toLowerCase()}</Text>
          </View>

          {checkboxField ? (
            <Text style={styles.checkboxHint}>{checkboxBehaviorHint(field.fieldName)}</Text>
          ) : null}
        </View>
      )}
    </Pressable>
  );
}

const styles = StyleSheet.create({
  card: {
    borderRadius: 14,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel,
    padding: spacing.md,
    gap: spacing.sm
  },
  cardActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  header: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  heading: {
    flex: 1,
    gap: spacing.xxs
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink,
    fontSize: 17
  },
  metaRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    flexWrap: 'wrap'
  },
  typeBadge: {
    ...typography.caption,
    color: colors.accent2,
    backgroundColor: colors.badgeBg,
    paddingHorizontal: spacing.sm,
    paddingVertical: 2,
    borderRadius: 999,
    overflow: 'hidden'
  },
  meta: {
    ...typography.caption,
    color: colors.muted
  },
  chevron: {
    ...typography.caption,
    color: colors.muted,
    marginTop: 4
  },
  collapsedFlags: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.sm
  },
  flagOn: {
    ...typography.caption,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  flagMuted: {
    ...typography.caption,
    color: colors.muted
  },
  body: {
    gap: spacing.sm,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border
  },
  toggleRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    minHeight: spacing.touchMin,
    gap: spacing.sm
  },
  toggleTitle: {
    ...typography.bodyStrong,
    color: colors.ink,
    flex: 1
  },
  techBlock: {
    gap: 4,
    padding: spacing.sm,
    borderRadius: 10,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border
  },
  techHeading: {
    ...typography.caption,
    color: colors.muted,
    marginBottom: 2
  },
  techLine: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'Courier'
  },
  checkboxHint: {
    ...typography.caption,
    color: colors.muted
  }
});
