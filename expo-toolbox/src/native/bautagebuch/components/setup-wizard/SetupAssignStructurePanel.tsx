import { Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { SingleLineText } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { getStructureItems } from '../../lib/setup-structure';
import {
  getWizardState,
  isFieldAssigned,
  resolveFieldDisplayLabel,
  type MappingField
} from '../../lib/setup-mapping';
import type { SetupStructureItem } from '../../types';

type Props = {
  setupModel: Record<string, unknown>;
  mappingFields: MappingField[];
  currentField: MappingField | null;
  readOnly?: boolean;
  onSelectGroup: (item: SetupStructureItem) => void;
  onSelectTable: (item: SetupStructureItem) => void;
};

function groupFields(
  setupModel: Record<string, unknown>,
  mappingFields: MappingField[],
  groupId: string
): string[] {
  const wizard = getWizardState(setupModel);
  return mappingFields
    .filter((field) => wizard.assignments[field.fieldId] === groupId)
    .map((field) => resolveFieldDisplayLabel(field, wizard));
}

function tableColumnFields(
  setupModel: Record<string, unknown>,
  mappingFields: MappingField[],
  tableId: string,
  columnId: string
): string[] {
  const wizard = getWizardState(setupModel);
  return mappingFields
    .filter(
      (field) =>
        wizard.tableAssignments[field.fieldId]?.tableId === tableId &&
        wizard.tableAssignments[field.fieldId]?.columnId === columnId
    )
    .map((field) => resolveFieldDisplayLabel(field, wizard));
}

export function SetupAssignStructurePanel({
  setupModel,
  mappingFields,
  currentField,
  readOnly = false,
  onSelectGroup,
  onSelectTable
}: Props) {
  const structure = getStructureItems(setupModel);
  const wizard = getWizardState(setupModel);
  const currentLabel = currentField ? resolveFieldDisplayLabel(currentField, wizard) : null;

  const handleSelect = (item: SetupStructureItem) => {
    if (readOnly || !currentField) return;
    void hapticSelection();
    if (item.type === 'group') onSelectGroup(item);
    else onSelectTable(item);
  };

  return (
    <ScrollView
      style={styles.scroll}
      contentContainerStyle={styles.content}
      keyboardShouldPersistTaps="handled"
      showsVerticalScrollIndicator={false}
    >
      <View style={styles.prompt}>
        <Text style={styles.heading}>Wohin gehört dieses Feld?</Text>
        {currentLabel ? (
          <View style={styles.currentField}>
            <Text style={styles.currentHint}>Erkanntes Feld:</Text>
            <Text style={styles.currentLabel}>„{currentLabel}"</Text>
          </View>
        ) : (
          <Text style={styles.doneCopy}>Alle Felder sind zugeordnet.</Text>
        )}
      </View>

      <View style={styles.list}>
        {structure.map((item) => {
          const isGroup = item.type === 'group';
          const icon = isGroup ? 'file-document-outline' : 'table';
          const tint = isGroup ? colors.accent2 : colors.info;
          const bg = isGroup ? 'rgba(36, 50, 64, 0.1)' : 'rgba(42, 95, 143, 0.12)';
          const assignedChildren = isGroup
            ? groupFields(setupModel, mappingFields, item.id)
            : [];

          return (
            <View key={item.id} style={styles.block}>
              <Pressable
                accessibilityRole="button"
                disabled={readOnly || !currentField}
                style={({ pressed }) => [
                  styles.target,
                  shadows.card,
                  pressed && currentField ? styles.targetPressed : null,
                  !currentField ? styles.targetDisabled : null
                ]}
                onPress={() => handleSelect(item)}
              >
                <View style={[styles.iconWrap, { backgroundColor: bg }]}>
                  <MaterialCommunityIcons name={icon} size={22} color={tint} />
                </View>
                <SingleLineText style={styles.targetLabel}>{item.name}</SingleLineText>
                <MaterialCommunityIcons name="chevron-right" size={20} color={colors.muted} />
              </Pressable>

              {isGroup && assignedChildren.length > 0 ? (
                <View style={styles.children}>
                  {assignedChildren.map((label) => (
                    <View key={`${item.id}_${label}`} style={styles.childRow}>
                      <Text style={styles.childBranch}>└</Text>
                      <SingleLineText style={styles.childLabel}>{label}</SingleLineText>
                    </View>
                  ))}
                </View>
              ) : null}

              {!isGroup && item.columns.length > 0 ? (
                <View style={styles.children}>
                  {item.columns.flatMap((column) => {
                    const labels = tableColumnFields(setupModel, mappingFields, item.id, column.id);
                    return labels.map((label) => (
                      <View key={`${item.id}_${column.id}_${label}`} style={styles.childRow}>
                        <Text style={styles.childBranch}>└</Text>
                        <SingleLineText style={styles.childLabel}>{label}</SingleLineText>
                      </View>
                    ));
                  })}
                </View>
              ) : null}
            </View>
          );
        })}
      </View>

      {currentField && isFieldAssigned(currentField.fieldId, wizard) ? (
        <Text style={styles.assignedNote}>Dieses Feld ist bereits zugeordnet.</Text>
      ) : null}
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  scroll: {
    flex: 1
  },
  content: {
    padding: spacing.pageX,
    paddingBottom: spacing.xl,
    gap: spacing.md
  },
  prompt: {
    gap: spacing.sm
  },
  heading: {
    ...typography.title,
    color: colors.ink
  },
  currentField: {
    gap: 4,
    padding: spacing.md,
    borderRadius: 14,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border
  },
  currentHint: {
    ...typography.caption,
    color: colors.muted
  },
  currentLabel: {
    ...typography.subtitle,
    color: colors.accent2
  },
  doneCopy: {
    ...typography.body,
    color: colors.muted,
    lineHeight: 22
  },
  list: {
    gap: spacing.sm
  },
  block: {
    gap: spacing.xxs
  },
  target: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: spacing.touchMin + 8,
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
    borderRadius: 14,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  targetPressed: {
    opacity: 0.92,
    transform: [{ scale: 0.995 }]
  },
  targetDisabled: {
    opacity: 0.65
  },
  iconWrap: {
    width: 40,
    height: 40,
    borderRadius: 12,
    alignItems: 'center',
    justifyContent: 'center'
  },
  targetLabel: {
    ...typography.bodyStrong,
    color: colors.ink,
    flex: 1
  },
  children: {
    paddingLeft: spacing.lg,
    gap: 4
  },
  childRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs
  },
  childBranch: {
    ...typography.caption,
    color: colors.muted,
    width: 14
  },
  childLabel: {
    ...typography.body,
    color: colors.ink,
    flex: 1
  },
  assignedNote: {
    ...typography.caption,
    color: colors.muted,
    textAlign: 'center'
  }
});
