import { useMemo } from 'react';
import { Pressable, StyleSheet, Switch, Text, View } from 'react-native';
import * as ImagePicker from 'expo-image-picker';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { PrimaryButton, TextField } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import {
  listCheckboxGroupOptions,
  listTableColumnsForMeta,
  normalizeSetupFieldType,
  resolveHybridFieldEditLabel,
  resolveTableMeta,
  setupFieldTypeLabel,
  updateFieldSettingsTarget,
  updateTableColumnMeta,
  updateTableMetaFlags,
  updateTableMetaLabel
} from '../../lib/setup-field-settings';
import { DEFAULT_WEATHER_METRIC, SETUP_WEATHER_METRIC_OPTIONS } from '../../lib/weather-metrics';
import type {
  DetectedField,
  FieldSettingsTarget,
  SetupFieldConfig,
  SetupFieldDateMode,
  SetupFieldType,
  SetupWeatherMetric
} from '../../types';
import type { MappingField } from '../../lib/setup-mapping';

type Props = {
  setupModel: Record<string, unknown>;
  target: FieldSettingsTarget;
  field: SetupFieldConfig | null;
  detectedFields: DetectedField[];
  mappingFields?: MappingField[];
  draftLabels?: Record<string, string>;
  readOnly?: boolean;
  onChange: (next: Record<string, unknown>) => void;
  onFieldLabelChange?: (fieldId: string, label: string) => void;
};

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

export function SetupFieldSettingsForm({
  setupModel,
  target,
  field,
  detectedFields,
  mappingFields = [],
  draftLabels = {},
  readOnly = false,
  onChange,
  onFieldLabelChange
}: Props) {
  const checkboxGroups = useMemo(
    () => listCheckboxGroupOptions(setupModel, detectedFields),
    [setupModel, detectedFields]
  );

  if (target.kind === 'table-meta') {
    return (
      <TableMetaForm
        setupModel={setupModel}
        tableId={target.tableId}
        readOnly={readOnly}
        onChange={onChange}
      />
    );
  }

  if (!field) {
    return (
      <View style={styles.empty}>
        <Text style={styles.emptyText}>Kein Feld ausgewählt.</Text>
      </View>
    );
  }

  const fieldType = normalizeSetupFieldType(field, detectedFields);

  const patchField = (patch: Partial<SetupFieldConfig>) => {
    onChange(updateFieldSettingsTarget(setupModel, target, patch));
  };

  const pickSignatureImage = async () => {
    const permission = await ImagePicker.requestMediaLibraryPermissionsAsync();
    if (!permission.granted) return;
    const result = await ImagePicker.launchImageLibraryAsync({
      mediaTypes: ['images'],
      quality: 0.9
    });
    if (result.canceled || !result.assets[0]?.uri) return;
    patchField({ signatureImageUri: result.assets[0].uri, signatureMode: 'image' });
  };

  return (
    <View style={styles.root}>
      {fieldType === 'static_text' ? (
        <TextField
          label="Textinhalt"
          value={field.staticText || ''}
          editable={!readOnly}
          onChangeText={(staticText) => patchField({ staticText })}
          multiline
        />
      ) : (
        <TextField
          label="Name des Feldes"
          value={resolveHybridFieldEditLabel(setupModel, target, field, mappingFields, draftLabels)}
          editable={!readOnly}
          onChangeText={(label) => {
            if (onFieldLabelChange && field.fieldId) {
              onFieldLabelChange(field.fieldId, label);
              return;
            }
            patchField({ label });
          }}
        />
      )}

      {fieldType === 'text' ? (
        <>
          <ToggleRow
            title="Pflichtfeld"
            value={Boolean(field.required)}
            disabled={readOnly}
            onValueChange={(required) => patchField({ required })}
          />
          <TextField
            label="Platzhalter"
            value={field.placeholder || field.hint || ''}
            editable={!readOnly}
            onChangeText={(placeholder) => patchField({ placeholder, hint: placeholder })}
            placeholder="Optional"
          />
          <ToggleRow
            title="Im BTB ausblenden"
            value={Boolean(field.skipped)}
            disabled={readOnly}
            onValueChange={(skipped) => patchField({ skipped })}
          />
          <ToggleRow
            title="Mehrzeilig"
            value={Boolean(field.multiline)}
            disabled={readOnly}
            onValueChange={(multiline) => patchField({ multiline })}
          />
        </>
      ) : null}

      {fieldType === 'number' ? (
        <>
          <ToggleRow
            title="Pflichtfeld"
            value={Boolean(field.required)}
            disabled={readOnly}
            onValueChange={(required) => patchField({ required })}
          />
          <ToggleRow
            title="Im BTB ausblenden"
            value={Boolean(field.skipped)}
            disabled={readOnly}
            onValueChange={(skipped) => patchField({ skipped })}
          />
        </>
      ) : null}

      {fieldType === 'datetime' ? (
        <>
          <ToggleRow
            title="Pflichtfeld"
            value={Boolean(field.required)}
            disabled={readOnly}
            onValueChange={(required) => patchField({ required })}
          />
          <ToggleRow
            title="Im BTB ausblenden"
            value={Boolean(field.skipped)}
            disabled={readOnly}
            onValueChange={(skipped) => patchField({ skipped })}
          />
          <ToggleRow
            title="Heutiges Datum übernehmen"
            value={Boolean(field.useCurrentDate)}
            disabled={readOnly}
            onValueChange={(useCurrentDate) => patchField({ useCurrentDate })}
          />
          <Text style={styles.sectionLabel}>Auswahl</Text>
          {(['date', 'time', 'datetime'] as SetupFieldDateMode[]).map((mode) => {
            const active = (field.dateMode || 'date') === mode;
            const label =
              mode === 'date' ? 'Datum' : mode === 'time' ? 'Uhrzeit' : 'Datum und Uhrzeit';
            return (
              <Pressable
                key={mode}
                style={[styles.choiceRow, active ? styles.choiceRowActive : null]}
                disabled={readOnly}
                onPress={() => {
                  void hapticSelection();
                  patchField({ dateMode: mode });
                }}
              >
                <MaterialCommunityIcons
                  name={active ? 'radiobox-marked' : 'radiobox-blank'}
                  size={20}
                  color={active ? colors.accent : colors.muted}
                />
                <Text style={active ? styles.choiceLabelActive : styles.choiceLabel}>{label}</Text>
              </Pressable>
            );
          })}
        </>
      ) : null}

      {fieldType === 'checkbox' ? (
        <>
          <ToggleRow
            title="Pflichtfeld"
            value={Boolean(field.required)}
            disabled={readOnly}
            onValueChange={(required) => patchField({ required })}
          />
          <ToggleRow
            title="Im BTB ausblenden"
            value={Boolean(field.skipped)}
            disabled={readOnly}
            onValueChange={(skipped) => patchField({ skipped })}
          />
          <ToggleRow
            title="Abhängigkeit mit anderer Checkbox"
            value={Boolean(field.checkboxExclusiveGroup)}
            disabled={readOnly}
            onValueChange={(enabled) =>
              patchField({
                checkboxExclusiveGroup: enabled
                  ? field.checkboxExclusiveGroup || resolveTargetGroupFallback(target)
                  : undefined
              })
            }
          />
          {field.checkboxExclusiveGroup ? (
            <View style={styles.groupPicker}>
              <Text style={styles.sectionLabel}>Checkbox-Gruppe</Text>
              {checkboxGroups.length === 0 ? (
                <Text style={styles.hint}>
                  Weise anderen Checkboxen dieselbe Gruppe zu, damit nur eine aktiv sein kann.
                </Text>
              ) : (
                checkboxGroups.map((group) => {
                  const active = field.checkboxExclusiveGroup === group.id;
                  return (
                    <Pressable
                      key={group.id}
                      style={[styles.choiceRow, active ? styles.choiceRowActive : null]}
                      disabled={readOnly}
                      onPress={() => patchField({ checkboxExclusiveGroup: group.id })}
                    >
                      <Text style={active ? styles.choiceLabelActive : styles.choiceLabel}>
                        {group.label}
                      </Text>
                    </Pressable>
                  );
                })
              )}
              <TextField
                label="Gruppen-ID"
                value={field.checkboxExclusiveGroup || ''}
                editable={!readOnly}
                onChangeText={(checkboxExclusiveGroup) => patchField({ checkboxExclusiveGroup })}
              />
            </View>
          ) : null}
        </>
      ) : null}

      {fieldType === 'select' ? (
        <>
          <ToggleRow
            title="Pflichtfeld"
            value={Boolean(field.required)}
            disabled={readOnly}
            onValueChange={(required) => patchField({ required })}
          />
          <ToggleRow
            title="Im BTB ausblenden"
            value={Boolean(field.skipped)}
            disabled={readOnly}
            onValueChange={(skipped) => patchField({ skipped })}
          />
          <TextField
            label="Auswahloptionen (kommagetrennt)"
            value={(field.options || []).join(', ')}
            editable={!readOnly}
            onChangeText={(raw) =>
              patchField({
                options: raw
                  .split(',')
                  .map((entry) => entry.trim())
                  .filter(Boolean)
              })
            }
            multiline
          />
        </>
      ) : null}

      {fieldType === 'weather' ? (
        <>
          <ToggleRow
            title="Pflichtfeld"
            value={Boolean(field.required)}
            disabled={readOnly}
            onValueChange={(required) => patchField({ required })}
          />
          <TextField
            label="Platzhalter"
            value={field.placeholder || field.hint || ''}
            editable={!readOnly}
            onChangeText={(placeholder) => patchField({ placeholder, hint: placeholder })}
            placeholder="Optional"
          />
          <ToggleRow
            title="Im BTB ausblenden"
            value={Boolean(field.skipped)}
            disabled={readOnly}
            onValueChange={(skipped) => patchField({ skipped })}
          />
          <Text style={styles.sectionLabel}>Wetterwert</Text>
          {SETUP_WEATHER_METRIC_OPTIONS.map((option) => {
            const active = (field.weatherMetric || DEFAULT_WEATHER_METRIC) === option.value;
            return (
              <Pressable
                key={option.value}
                style={[styles.choiceRow, active ? styles.choiceRowActive : null]}
                disabled={readOnly}
                onPress={() => {
                  void hapticSelection();
                  patchField({ weatherMetric: option.value as SetupWeatherMetric });
                }}
              >
                <MaterialCommunityIcons
                  name={active ? 'radiobox-marked' : 'radiobox-blank'}
                  size={20}
                  color={active ? colors.accent : colors.muted}
                />
                <Text style={active ? styles.choiceLabelActive : styles.choiceLabel}>{option.label}</Text>
              </Pressable>
            );
          })}
        </>
      ) : null}

      {fieldType === 'signature' ? (
        <>
          <ToggleRow
            title="Pflichtfeld"
            value={Boolean(field.required)}
            disabled={readOnly}
            onValueChange={(required) => patchField({ required })}
          />
          <ToggleRow
            title="Im BTB ausblenden"
            value={Boolean(field.skipped)}
            disabled={readOnly}
            onValueChange={(skipped) => patchField({ skipped })}
          />
          <Text style={styles.sectionLabel}>Unterschrift einfügen</Text>
          {(['draw', 'image'] as const).map((mode) => {
            const active = (field.signatureMode || 'draw') === mode;
            const label = mode === 'draw' ? 'Direkt unterschreiben' : 'Unterschrift als Bild hinterlegen';
            return (
              <Pressable
                key={mode}
                style={[styles.choiceRow, active ? styles.choiceRowActive : null]}
                disabled={readOnly}
                onPress={() => {
                  void hapticSelection();
                  patchField({ signatureMode: mode });
                }}
              >
                <MaterialCommunityIcons
                  name={active ? 'radiobox-marked' : 'radiobox-blank'}
                  size={20}
                  color={active ? colors.accent : colors.muted}
                />
                <Text style={active ? styles.choiceLabelActive : styles.choiceLabel}>{label}</Text>
              </Pressable>
            );
          })}
          {field.signatureMode === 'image' ? (
            <PrimaryButton
              compact
              label={field.signatureImageUri ? 'Bild erneut wählen' : 'Bild auswählen'}
              variant="secondary"
              disabled={readOnly}
              onPress={() => void pickSignatureImage()}
            />
          ) : null}
        </>
      ) : null}

      {fieldType === 'table' && target.kind === 'table-cell' ? (
        <>
          <ToggleRow
            title="Pflichtfeld"
            value={Boolean(field.required)}
            disabled={readOnly}
            onValueChange={(required) => patchField({ required })}
          />
          <ToggleRow
            title="Im BTB ausblenden"
            value={Boolean(field.skipped)}
            disabled={readOnly}
            onValueChange={(skipped) => patchField({ skipped })}
          />
        </>
      ) : null}
    </View>
  );
}

function resolveTargetGroupFallback(target: FieldSettingsTarget): string {
  if (target.kind === 'single') return target.sectionId;
  if (target.kind === 'table-cell') return target.tableId;
  return 'checkbox-group';
}

function TableMetaForm({
  setupModel,
  tableId,
  readOnly,
  onChange
}: {
  setupModel: Record<string, unknown>;
  tableId: string;
  readOnly?: boolean;
  onChange: (next: Record<string, unknown>) => void;
}) {
  const meta = resolveTableMeta(setupModel, tableId);
  const columns = listTableColumnsForMeta(setupModel, tableId);
  if (!meta) return null;

  return (
    <View style={styles.root}>
      <TextField
        label="Name der Tabelle"
        value={meta.label}
        editable={!readOnly}
        onChangeText={(label) => onChange(updateTableMetaLabel(setupModel, tableId, label))}
      />
      <ToggleRow
        title="Im BTB ausblenden"
        value={Boolean(meta.skipped)}
        disabled={readOnly}
        onValueChange={(skipped) => onChange(updateTableMetaFlags(setupModel, tableId, { skipped }))}
      />
      <ToggleRow
        title="Mehrzeilig"
        value={Boolean(meta.multiline)}
        disabled={readOnly}
        onValueChange={(multiline) => onChange(updateTableMetaFlags(setupModel, tableId, { multiline }))}
      />

      <Text style={styles.sectionLabel}>Spalten verwalten</Text>
      {columns.map((column) => (
        <View key={column.columnId} style={styles.columnCard}>
          <TextField
            label="Name"
            value={column.label}
            editable={!readOnly}
            onChangeText={(label) =>
              onChange(updateTableColumnMeta(setupModel, tableId, column.columnId, { label }))
            }
          />
          <Text style={styles.columnType}>Feldtyp: {setupFieldTypeLabel(normalizeColumnType(column.type))}</Text>
          <ToggleRow
            title="Pflichtfeld"
            value={Boolean(column.required)}
            disabled={readOnly}
            onValueChange={(required) =>
              onChange(updateTableColumnMeta(setupModel, tableId, column.columnId, { required }))
            }
          />
          <ToggleRow
            title="Im BTB ausblenden"
            value={Boolean(column.skipped)}
            disabled={readOnly}
            onValueChange={(skipped) =>
              onChange(updateTableColumnMeta(setupModel, tableId, column.columnId, { skipped }))
            }
          />
        </View>
      ))}
    </View>
  );
}

function normalizeColumnType(type: string): SetupFieldType {
  if (type === 'checkbox') return 'checkbox';
  if (type === 'dropdown' || type === 'radio') return 'select';
  return 'text';
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.sm
  },
  empty: {
    padding: spacing.md,
    alignItems: 'center'
  },
  emptyText: {
    ...typography.body,
    color: colors.muted
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
  sectionLabel: {
    ...typography.label,
    color: colors.muted,
    marginTop: spacing.xxs
  },
  choiceRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: 44,
    paddingHorizontal: spacing.sm,
    borderRadius: 10,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  choiceRowActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  choiceLabel: {
    ...typography.body,
    color: colors.ink,
    flex: 1
  },
  choiceLabelActive: {
    ...typography.bodyStrong,
    color: colors.accent2,
    flex: 1
  },
  groupPicker: {
    gap: spacing.xs
  },
  hint: {
    ...typography.caption,
    color: colors.muted,
    lineHeight: 18
  },
  columnCard: {
    gap: spacing.xs,
    padding: spacing.md,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated
  },
  columnType: {
    ...typography.caption,
    color: colors.muted
  }
});
