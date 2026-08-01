import { Alert, Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { PrimaryButton, TextField } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { SETUP_FIELD_TYPE_OPTIONS } from '../../lib/setup-field-settings';
import {
  getWizardState,
  resolveFieldAssignmentSummary,
  resolveFieldDisplayLabel,
  resolveFieldEditLabel,
  type MappingField
} from '../../lib/setup-mapping';
import { getStructureItems } from '../../lib/setup-structure';
import { fieldSourceLabel, fieldSourceTone } from '../../lib/template-field';
import type { SetupFieldType, SetupStructureItem } from '../../types';
import { SetupFieldStepNavigator } from './SetupFieldStepNavigator';
import { SetupScrollView } from './SetupScrollView';

function mappingTypeLabel(type: string): string {
  const normalized =
    type === 'dropdown' || type === 'radio' ? 'select' : (type as SetupFieldType);
  return (
    SETUP_FIELD_TYPE_OPTIONS.find((option) => option.value === normalized)?.label || 'Textfeld'
  );
}

type Props = {
  mappingFields: MappingField[];
  setupModel: Record<string, unknown>;
  currentField: MappingField | null;
  currentFieldIndex: number;
  unassignedCount: number;
  draftLabels?: Record<string, string>;
  readOnly?: boolean;
  onSelectField: (index: number) => void;
  onShowInPdf: () => void;
  onAssignGroup: (item: SetupStructureItem) => void;
  onAssignTable: (item: SetupStructureItem) => void;
  onChangeFieldName: (fieldId: string, name: string) => void;
  onChangeFieldType: (fieldId: string, type: SetupFieldType) => void;
  onDeleteField: (fieldId: string) => void;
  bottomInset?: number;
};

export function SetupAssignFieldListPanel({
  mappingFields,
  setupModel,
  currentField,
  currentFieldIndex,
  unassignedCount,
  draftLabels = {},
  readOnly = false,
  onSelectField,
  onShowInPdf,
  onAssignGroup,
  onAssignTable,
  onChangeFieldName,
  onChangeFieldType,
  onDeleteField,
  bottomInset = 0
}: Props) {
  const wizard = getWizardState(setupModel);
  const structureItems = getStructureItems(setupModel);
  const groups = structureItems.filter((item) => item.type === 'group');
  const tables = structureItems.filter((item) => item.type === 'table');
  const resolveLabel = (field: MappingField) =>
    resolveFieldDisplayLabel(field, wizard, draftLabels);
  const assignment = currentField
    ? resolveFieldAssignmentSummary(setupModel, currentField.fieldId)
    : null;
  const fieldNumber = currentField?.displayOrder ?? currentFieldIndex + 1;
  const displayName = currentField ? resolveLabel(currentField) : null;

  const confirmDelete = () => {
    if (!currentField || readOnly) return;
    Alert.alert(
      'Feld entfernen',
      'Möchten Sie dieses Feld wirklich entfernen? Das Feld wird aus dieser Vorlage entfernt.',
      [
        { text: 'Abbrechen', style: 'cancel' },
        {
          text: 'Entfernen',
          style: 'destructive',
          onPress: () => onDeleteField(currentField.fieldId)
        }
      ]
    );
  };

  const goToField = (index: number) => {
    if (index < 0 || index >= mappingFields.length) return;
    onSelectField(index);
  };

  return (
    <SetupScrollView
      style={styles.scroll}
      contentContainerStyle={[styles.scrollContent, { paddingBottom: spacing.lg + bottomInset }]}
      showsVerticalScrollIndicator={false}
      nestedScrollEnabled
    >
      {mappingFields.length === 0 ? (
        <View style={styles.empty}>
          <Text style={styles.emptyTitle}>Noch keine Felder</Text>
          <Text style={styles.emptyCopy}>
            Wechseln Sie zur PDF-Ansicht und markieren Sie Bereiche mit „+ Feld markieren“.
          </Text>
        </View>
      ) : null}

      {currentField ? (
        <View style={styles.detail}>
          <SetupFieldStepNavigator
            fieldNumber={fieldNumber}
            total={mappingFields.length}
            canGoPrevious={currentFieldIndex > 0}
            canGoNext={currentFieldIndex < mappingFields.length - 1}
            onPrevious={() => goToField(currentFieldIndex - 1)}
            onNext={() => goToField(currentFieldIndex + 1)}
          />

          {displayName ? <Text style={styles.fieldName}>{displayName}</Text> : null}

          {unassignedCount > 0 ? (
            <Text style={styles.unassignedSummary}>
              Noch nicht zugeordnet:{' '}
              {unassignedCount} Feld{unassignedCount === 1 ? '' : 'er'}
            </Text>
          ) : null}

          <TextField
            label="Name"
            value={resolveFieldEditLabel(currentField, wizard, draftLabels)}
            onChangeText={(value) => onChangeFieldName(currentField.fieldId, value)}
            editable={!readOnly}
            placeholder="Feldname"
          />

          <Text style={styles.detailLabel}>Feldtyp</Text>
          <SetupScrollView horizontal showsHorizontalScrollIndicator={false} style={styles.typeRow} nestedScrollEnabled>
            {SETUP_FIELD_TYPE_OPTIONS.filter((option) => option.value !== 'table').map((option) => {
              const normalized =
                currentField.type === 'dropdown' || currentField.type === 'radio'
                  ? 'select'
                  : currentField.type;
              const active = option.value === normalized;
              return (
                <Pressable
                  key={option.value}
                  style={[styles.typeChip, active ? styles.typeChipActive : null]}
                  disabled={readOnly}
                  onPress={() => onChangeFieldType(currentField.fieldId, option.value)}
                >
                  <Text style={active ? styles.typeChipLabelActive : styles.typeChipLabel}>
                    {option.label}
                  </Text>
                </Pressable>
              );
            })}
          </SetupScrollView>

          <Text style={styles.detailLabel}>Quelle</Text>
          <Text
            style={[
              styles.detailValue,
              fieldSourceTone(currentField.source) === 'success'
                ? styles.sourceAuto
                : fieldSourceTone(currentField.source) === 'warning'
                  ? styles.sourceManual
                  : null
            ]}
          >
            {fieldSourceLabel(currentField.source)} · {mappingTypeLabel(currentField.type)}
          </Text>

          <Text style={styles.detailLabel}>Zuordnung</Text>
          {assignment?.kind === 'none' ? (
            <Text style={styles.unassigned}>Keine Gruppe zugewiesen</Text>
          ) : (
            <Text style={styles.detailValue}>
              {assignment?.kind === 'table' ? 'Tabelle' : 'Gruppe'}: {assignment?.label}
            </Text>
          )}

          {!readOnly && assignment?.kind === 'none' ? (
            <View style={styles.assignSection}>
              {groups.length > 0 ? (
                <View style={styles.assignBlock}>
                  <Text style={styles.assignHeading}>Gruppe</Text>
                  {groups.map((item) => (
                    <Pressable
                      key={item.id}
                      accessibilityRole="button"
                      style={styles.assignRow}
                      onPress={() => onAssignGroup(item)}
                    >
                      <MaterialCommunityIcons name="circle-outline" size={18} color={colors.accent} />
                      <Text style={styles.assignRowLabel}>{item.name}</Text>
                    </Pressable>
                  ))}
                </View>
              ) : null}
              {tables.length > 0 ? (
                <View style={styles.assignBlock}>
                  <Text style={styles.assignHeading}>Tabellen</Text>
                  {tables.map((item) => (
                    <Pressable
                      key={item.id}
                      accessibilityRole="button"
                      style={styles.assignRow}
                      onPress={() => onAssignTable(item)}
                    >
                      <MaterialCommunityIcons name="circle-outline" size={18} color={colors.accent} />
                      <Text style={styles.assignRowLabel}>{item.name}</Text>
                      {item.columns.length > 0 ? (
                        <Text style={styles.assignRowMeta}>
                          {item.columns.length} Spalte{item.columns.length === 1 ? '' : 'n'}
                        </Text>
                      ) : (
                        <Text style={styles.assignRowMeta}>Keine Spalten</Text>
                      )}
                    </Pressable>
                  ))}
                </View>
              ) : null}
            </View>
          ) : null}

          <View style={styles.detailActions}>
            <PrimaryButton label="In PDF anzeigen" variant="ghost" compact onPress={onShowInPdf} />
            {!readOnly ? (
              <Pressable accessibilityRole="button" style={styles.deleteBtn} onPress={confirmDelete}>
                <MaterialCommunityIcons name="trash-can-outline" size={18} color={colors.danger} />
                <Text style={styles.deleteLabel}>Feld entfernen</Text>
              </Pressable>
            ) : null}
          </View>
        </View>
      ) : null}
    </SetupScrollView>
  );
}

const styles = StyleSheet.create({
  scroll: {
    flex: 1,
    minHeight: 0
  },
  scrollContent: {
    padding: spacing.pageX,
    paddingBottom: spacing.lg,
    flexGrow: 1
  },
  empty: {
    padding: spacing.md,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    gap: spacing.xxs
  },
  emptyTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  emptyCopy: {
    ...typography.caption,
    color: colors.muted,
    lineHeight: 18
  },
  detail: {
    gap: spacing.xs
  },
  fieldName: {
    ...typography.title,
    color: colors.ink,
    marginBottom: spacing.xxs
  },
  unassignedSummary: {
    ...typography.body,
    color: colors.muted,
    marginBottom: spacing.xs
  },
  detailLabel: {
    ...typography.label,
    color: colors.muted,
    marginTop: spacing.xxs
  },
  detailValue: {
    ...typography.body,
    color: colors.ink
  },
  sourceAuto: {
    color: colors.success
  },
  sourceManual: {
    color: colors.warning
  },
  typeRow: {
    flexGrow: 0
  },
  typeChip: {
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.xs,
    borderRadius: 999,
    borderWidth: 1,
    borderColor: colors.border,
    marginRight: spacing.xxs,
    backgroundColor: colors.panelElevated
  },
  typeChipActive: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  typeChipLabel: {
    ...typography.caption,
    color: colors.muted
  },
  typeChipLabelActive: {
    ...typography.caption,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  unassigned: {
    ...typography.body,
    color: colors.muted
  },
  assignSection: {
    gap: spacing.sm,
    marginTop: spacing.xxs
  },
  assignBlock: {
    gap: spacing.xxs
  },
  assignHeading: {
    ...typography.label,
    color: colors.muted
  },
  assignRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: 44,
    paddingHorizontal: spacing.sm,
    borderRadius: 10,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated
  },
  assignRowLabel: {
    ...typography.body,
    color: colors.ink,
    flex: 1
  },
  assignRowMeta: {
    ...typography.caption,
    color: colors.muted
  },
  detailActions: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginTop: spacing.sm,
    gap: spacing.sm
  },
  deleteBtn: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs,
    minHeight: 36,
    paddingHorizontal: spacing.xs
  },
  deleteLabel: {
    ...typography.caption,
    color: colors.danger,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  }
});
