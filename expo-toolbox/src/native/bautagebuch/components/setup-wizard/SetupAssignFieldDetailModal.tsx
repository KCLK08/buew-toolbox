import { useEffect, useMemo, useState } from 'react';
import { Modal, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { PrimaryButton, TextField } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { SETUP_FIELD_TYPE_OPTIONS } from '../../lib/setup-field-settings';
import {
  canDeleteMappingField,
  canEditMappingFieldGeometry,
  isManualMappingField
} from '../../lib/setup-manual-field';
import {
  getWizardState,
  isFieldAssigned,
  resolveFieldAssignmentSummary,
  resolveFieldDisplayLabel,
  resolveFieldEditLabel,
  type MappingField
} from '../../lib/setup-mapping';
import { getStructureItems } from '../../lib/setup-structure';
import { fieldSourceLabel, fieldSourceTone } from '../../lib/template-field';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import type { SetupFieldType, SetupStructureItem } from '../../types';
import { SetupModalKeyboardFrame } from './SetupModalKeyboardFrame';
import { SetupScrollView } from './SetupScrollView';

function mappingTypeLabel(type: string): string {
  const normalized =
    type === 'dropdown' || type === 'radio' ? 'select' : (type as SetupFieldType);
  return (
    SETUP_FIELD_TYPE_OPTIONS.find((option) => option.value === normalized)?.label || 'Textfeld'
  );
}

function assignmentStatusLabel(
  assigned: boolean,
  assignment: ReturnType<typeof resolveFieldAssignmentSummary>
): string {
  if (!assigned || assignment.kind === 'none') return 'Nicht zugeordnet';
  if (assignment.kind === 'table') return `Tabelle: ${assignment.label}`;
  return `Gruppe: ${assignment.label}`;
}

type Props = {
  visible: boolean;
  field: MappingField | null;
  setupModel: Record<string, unknown>;
  draftLabels?: Record<string, string>;
  readOnly?: boolean;
  onClose: () => void;
  onChangeFieldName: (fieldId: string, name: string) => void;
  onChangeFieldType: (fieldId: string, type: SetupFieldType) => void;
  onAssignGroup: (item: SetupStructureItem) => void;
  onAssignTable: (item: SetupStructureItem) => void;
  onClearAssignment: (fieldId: string) => void;
  onEditPosition?: (fieldId: string) => void;
  onDelete?: (fieldId: string) => void;
};

export function SetupAssignFieldDetailModal({
  visible,
  field,
  setupModel,
  draftLabels = {},
  readOnly = false,
  onClose,
  onChangeFieldName,
  onChangeFieldType,
  onAssignGroup,
  onAssignTable,
  onClearAssignment,
  onEditPosition,
  onDelete
}: Props) {
  const insets = useSafeAreaInsets();
  const wizard = useMemo(() => getWizardState(setupModel), [setupModel]);
  const structureItems = useMemo(() => getStructureItems(setupModel), [setupModel]);
  const groups = structureItems.filter((item) => item.type === 'group');
  const tables = structureItems.filter((item) => item.type === 'table');
  const [showReassign, setShowReassign] = useState(false);

  useEffect(() => {
    if (visible) {
      setShowReassign(false);
    }
  }, [visible, field?.fieldId]);

  if (!field) return null;

  const assignment = resolveFieldAssignmentSummary(setupModel, field.fieldId);
  const assigned = isFieldAssigned(field.fieldId, wizard);
  const isManual = isManualMappingField(field);
  const canDelete = canDeleteMappingField(field);
  const canEditGeometry = canEditMappingFieldGeometry(field);
  const displayName = resolveFieldDisplayLabel(field, wizard, draftLabels);
  const sourceTone = fieldSourceTone(field.source);
  const normalizedType =
    field.type === 'dropdown' || field.type === 'radio' ? 'select' : (field.type as SetupFieldType);

  const handleAssignGroup = (item: SetupStructureItem) => {
    void hapticSelection();
    onAssignGroup(item);
    onClose();
  };

  const handleAssignTable = (item: SetupStructureItem) => {
    void hapticSelection();
    onAssignTable(item);
    onClose();
  };

  return (
    <Modal visible={visible} animationType="slide" transparent onRequestClose={onClose}>
      <Pressable style={styles.backdrop} onPress={onClose} accessibilityRole="button" />
      <SetupModalKeyboardFrame>
        <View style={[styles.sheet, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
          <View style={styles.handle} />
          <View style={styles.header}>
            <Text style={styles.title}>Feld</Text>
            <Pressable accessibilityRole="button" onPress={onClose} hitSlop={12}>
              <MaterialCommunityIcons name="close" size={22} color={colors.muted} />
            </Pressable>
          </View>

          <SetupScrollView
            style={styles.body}
            contentContainerStyle={styles.bodyContent}
            keyboardShouldPersistTaps="handled"
          >
            <Text style={styles.fieldName}>{displayName || 'Unbenanntes Feld'}</Text>

            <TextField
              label="Feldname"
              value={resolveFieldEditLabel(field, wizard, draftLabels)}
              onChangeText={(value) => onChangeFieldName(field.fieldId, value)}
              editable={!readOnly}
              placeholder="Feldname"
            />

            <Text style={styles.label}>Feldtyp</Text>
            <ScrollView horizontal showsHorizontalScrollIndicator={false} style={styles.typeRow}>
              {SETUP_FIELD_TYPE_OPTIONS.filter((option) => option.value !== 'table').map((option) => {
                const active = option.value === normalizedType;
                return (
                  <Pressable
                    key={option.value}
                    style={[styles.typeChip, active ? styles.typeChipActive : null]}
                    disabled={readOnly}
                    onPress={() => onChangeFieldType(field.fieldId, option.value)}
                  >
                    <Text style={active ? styles.typeChipLabelActive : styles.typeChipLabel}>
                      {option.label}
                    </Text>
                  </Pressable>
                );
              })}
            </ScrollView>

            <Text style={styles.label}>Quelle</Text>
            <Text
              style={[
                styles.value,
                sourceTone === 'success'
                  ? styles.sourceAuto
                  : sourceTone === 'warning'
                    ? styles.sourceManual
                    : null
              ]}
            >
              {fieldSourceLabel(field.source)} · {mappingTypeLabel(field.type)}
            </Text>

            <Text style={styles.label}>Zuordnung</Text>
            <Text style={styles.value}>{assignmentStatusLabel(assigned, assignment)}</Text>

            {!readOnly ? (
              <View style={styles.actionBlock}>
                <PrimaryButton
                  label={showReassign ? 'Zuordnung schließen' : 'Zuordnung ändern'}
                  variant="secondary"
                  compact
                  onPress={() => setShowReassign((current) => !current)}
                />
                {assigned ? (
                  <Pressable
                    accessibilityRole="button"
                    style={styles.clearBtn}
                    onPress={() => {
                      void hapticSelection();
                      onClearAssignment(field.fieldId);
                      onClose();
                    }}
                  >
                    <Text style={styles.clearLabel}>Zuordnung entfernen</Text>
                  </Pressable>
                ) : null}
              </View>
            ) : null}

            {!readOnly && showReassign ? (
              <View style={styles.assignSection}>
                {groups.length > 0 ? (
                  <View style={styles.assignBlock}>
                    <Text style={styles.assignHeading}>Gruppe</Text>
                    {groups.map((item) => (
                      <Pressable
                        key={item.id}
                        accessibilityRole="button"
                        style={styles.assignRow}
                        onPress={() => handleAssignGroup(item)}
                      >
                        <MaterialCommunityIcons name="folder-outline" size={18} color={colors.accent} />
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
                        onPress={() => handleAssignTable(item)}
                      >
                        <MaterialCommunityIcons name="table" size={18} color={colors.accent} />
                        <Text style={styles.assignRowLabel}>{item.name}</Text>
                      </Pressable>
                    ))}
                  </View>
                ) : null}
              </View>
            ) : null}

            {!readOnly && isManual ? (
              <View style={styles.manualActions}>
                {canEditGeometry && onEditPosition ? (
                  <Pressable
                    accessibilityRole="button"
                    style={styles.secondaryAction}
                    onPress={() => {
                      onEditPosition(field.fieldId);
                      onClose();
                    }}
                  >
                    <MaterialCommunityIcons name="vector-square-edit" size={18} color={colors.accent} />
                    <Text style={styles.secondaryActionLabel}>Position ändern</Text>
                  </Pressable>
                ) : null}
                {canDelete && onDelete ? (
                  <Pressable
                    accessibilityRole="button"
                    style={styles.deleteAction}
                    onPress={() => onDelete(field.fieldId)}
                  >
                    <MaterialCommunityIcons name="trash-can-outline" size={18} color={colors.danger} />
                    <Text style={styles.deleteLabel}>Feld löschen</Text>
                  </Pressable>
                ) : null}
              </View>
            ) : null}
          </SetupScrollView>
        </View>
      </SetupModalKeyboardFrame>
    </Modal>
  );
}

const styles = StyleSheet.create({
  backdrop: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: 'rgba(0, 0, 0, 0.35)'
  },
  sheet: {
    marginTop: 'auto',
    maxHeight: '88%',
    borderTopLeftRadius: 20,
    borderTopRightRadius: 20,
    backgroundColor: colors.panel,
    borderTopWidth: 1,
    borderColor: colors.border
  },
  handle: {
    alignSelf: 'center',
    width: 40,
    height: 4,
    borderRadius: 999,
    backgroundColor: colors.border,
    marginTop: spacing.xs,
    marginBottom: spacing.xxs
  },
  header: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingHorizontal: spacing.pageX,
    paddingBottom: spacing.xs
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  body: {
    maxHeight: 520
  },
  bodyContent: {
    paddingHorizontal: spacing.pageX,
    paddingBottom: spacing.md,
    gap: spacing.xs
  },
  fieldName: {
    ...typography.title,
    color: colors.ink,
    marginBottom: spacing.xxs
  },
  label: {
    ...typography.label,
    color: colors.muted,
    marginTop: spacing.xxs
  },
  value: {
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
  actionBlock: {
    gap: spacing.xs,
    marginTop: spacing.xs
  },
  clearBtn: {
    alignSelf: 'flex-start',
    minHeight: 36,
    justifyContent: 'center',
    paddingHorizontal: spacing.xs
  },
  clearLabel: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  assignSection: {
    gap: spacing.sm,
    marginTop: spacing.xs
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
  manualActions: {
    gap: spacing.xs,
    marginTop: spacing.sm,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border
  },
  secondaryAction: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs,
    minHeight: 40
  },
  secondaryActionLabel: {
    ...typography.body,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  deleteAction: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs,
    minHeight: 40
  },
  deleteLabel: {
    ...typography.body,
    color: colors.danger,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  }
});
