import { Alert, Pressable, ScrollView, StyleSheet, Text, TextInput, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { SETUP_FIELD_TYPE_OPTIONS } from '../../lib/setup-field-settings';
import {
  getWizardState,
  isFieldAssigned,
  resolveFieldAssignmentSummary,
  resolveFieldDisplayLabel,
  type MappingField
} from '../../lib/setup-mapping';
import { getStructureItems } from '../../lib/setup-structure';
import { fieldHasGeometry, fieldSourceLabel, fieldSourceTone } from '../../lib/template-field';
import type { SetupFieldType, SetupStructureItem } from '../../types';

function mappingTypeLabel(type: string): string {
  const normalized =
    type === 'dropdown' || type === 'radio' ? 'select' : (type as SetupFieldType);
  return (
    SETUP_FIELD_TYPE_OPTIONS.find((option) => option.value === normalized)?.label || 'Textfeld'
  );
}

function sourceDotColor(source: MappingField['source']): string {
  const tone = fieldSourceTone(source);
  if (tone === 'warning') return colors.warning;
  if (tone === 'neutral') return colors.muted;
  return colors.success;
}

type Props = {
  mappingFields: MappingField[];
  setupModel: Record<string, unknown>;
  currentField: MappingField | null;
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

  return (
    <View style={styles.root}>
      <ScrollView
        style={styles.listScroll}
        contentContainerStyle={[styles.listContent, { paddingBottom: spacing.md + bottomInset }]}
        keyboardShouldPersistTaps="handled"
        showsVerticalScrollIndicator={false}
      >
        <Text style={styles.heading}>Felder</Text>
        <Text style={styles.subheading}>
          {mappingFields.length} Feld{mappingFields.length === 1 ? '' : 'er'} · Prüfen, bearbeiten und
          zuordnen
        </Text>

        {mappingFields.length === 0 ? (
          <View style={styles.empty}>
            <Text style={styles.emptyTitle}>Noch keine Felder</Text>
            <Text style={styles.emptyCopy}>
              Wechseln Sie zur PDF-Ansicht und markieren Sie Bereiche mit „+ Feld markieren“.
            </Text>
          </View>
        ) : null}

        {mappingFields.map((field) => {
          const label = resolveLabel(field);
          const assigned = isFieldAssigned(field.fieldId, wizard);
          const fieldAssignment = resolveFieldAssignmentSummary(setupModel, field.fieldId);
          const isCurrent = field.fieldId === currentField?.fieldId;
          const hasGeometry = fieldHasGeometry(field);
          return (
            <Pressable
              key={field.fieldId}
              accessibilityRole="button"
              style={[styles.card, isCurrent ? styles.cardCurrent : null]}
              onPress={() => {
                void hapticSelection();
                const index = mappingFields.findIndex((entry) => entry.fieldId === field.fieldId);
                if (index >= 0) onSelectField(index);
              }}
            >
              <View style={styles.cardHeader}>
                <View style={[styles.sourceDot, { backgroundColor: sourceDotColor(field.source) }]} />
                <View style={styles.cardCopy}>
                  <Text style={styles.cardTitle}>{label}</Text>
                  <Text style={styles.cardMeta}>
                    {mappingTypeLabel(field.type)} · {fieldSourceLabel(field.source)}
                  </Text>
                  <Text style={styles.cardMeta}>
                    {assigned
                      ? fieldAssignment.kind === 'table'
                        ? `Tabelle: ${fieldAssignment.label}`
                        : `Gruppe: ${fieldAssignment.label}`
                      : 'Keine Gruppe zugewiesen'}
                    {hasGeometry ? ` · Seite ${field.page}` : ''}
                  </Text>
                </View>
                <Text style={styles.cardIndex}>{field.displayOrder}</Text>
              </View>
            </Pressable>
          );
        })}
      </ScrollView>

      {currentField ? (
        <View style={[styles.detail, { paddingBottom: spacing.lg + bottomInset }]}>
          <Text style={styles.detailTitle}>Felddetails</Text>

          <Text style={styles.detailLabel}>Name</Text>
          <TextInput
            value={resolveLabel(currentField)}
            onChangeText={(value) => onChangeFieldName(currentField.fieldId, value)}
            style={styles.input}
            editable={!readOnly}
            placeholder="Feldname"
          />

          <Text style={styles.detailLabel}>Feldtyp</Text>
          <ScrollView horizontal showsHorizontalScrollIndicator={false} style={styles.typeRow}>
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
          </ScrollView>

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
            {fieldSourceLabel(currentField.source)}
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
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    minHeight: 0
  },
  listScroll: {
    flex: 1
  },
  listContent: {
    padding: spacing.pageX,
    paddingBottom: spacing.md,
    gap: spacing.xs
  },
  heading: {
    ...typography.title,
    color: colors.ink
  },
  subheading: {
    ...typography.caption,
    color: colors.muted,
    marginBottom: spacing.xs
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
  card: {
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel,
    padding: spacing.sm
  },
  cardCurrent: {
    borderColor: colors.warning,
    backgroundColor: 'rgba(196, 140, 40, 0.06)'
  },
  cardHeader: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.sm
  },
  sourceDot: {
    width: 10,
    height: 10,
    borderRadius: 999,
    marginTop: 6
  },
  cardCopy: {
    flex: 1,
    gap: 2
  },
  cardTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  cardMeta: {
    ...typography.caption,
    color: colors.muted
  },
  cardIndex: {
    ...typography.caption,
    color: colors.muted
  },
  detail: {
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel,
    padding: spacing.pageX,
    paddingBottom: spacing.lg,
    gap: spacing.xs
  },
  detailTitle: {
    ...typography.subtitle,
    color: colors.ink
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
  input: {
    ...typography.body,
    color: colors.ink,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: 12,
    paddingHorizontal: spacing.md,
    minHeight: spacing.touchMin,
    backgroundColor: colors.panelElevated
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
