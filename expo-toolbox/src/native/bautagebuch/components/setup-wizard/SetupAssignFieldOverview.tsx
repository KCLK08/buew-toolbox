import { Modal, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { SingleLineText } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import {
  getWizardState,
  isFieldAssigned,
  resolveFieldDisplayLabel,
  type MappingField
} from '../../lib/setup-mapping';

type FieldStatus = 'assigned' | 'current' | 'open';

type Props = {
  visible: boolean;
  mappingFields: MappingField[];
  setupModel: Record<string, unknown>;
  currentFieldId: string | null;
  onClose: () => void;
  onSelectField: (index: number) => void;
};

function fieldStatus(
  field: MappingField,
  wizard: ReturnType<typeof getWizardState>,
  currentFieldId: string | null
): FieldStatus {
  if (field.fieldId === currentFieldId) return 'current';
  if (isFieldAssigned(field.fieldId, wizard)) return 'assigned';
  return 'open';
}

function statusIcon(status: FieldStatus): keyof typeof MaterialCommunityIcons.glyphMap {
  if (status === 'assigned') return 'circle';
  if (status === 'current') return 'circle';
  return 'circle-outline';
}

function statusColor(status: FieldStatus): string {
  if (status === 'assigned') return colors.success;
  if (status === 'current') return colors.warning;
  return colors.muted;
}

export function SetupAssignFieldOverview({
  visible,
  mappingFields,
  setupModel,
  currentFieldId,
  onClose,
  onSelectField
}: Props) {
  const insets = useSafeAreaInsets();
  const wizard = getWizardState(setupModel);

  return (
    <Modal visible={visible} animationType="slide" onRequestClose={onClose}>
      <View style={[styles.root, { paddingTop: insets.top }]}>
        <View style={styles.header}>
          <Pressable accessibilityRole="button" style={styles.headerBtn} onPress={onClose}>
            <Text style={styles.cancel}>Schließen</Text>
          </Pressable>
          <Text style={styles.title}>Felder</Text>
          <View style={styles.headerBtn} />
        </View>

        <View style={styles.legend}>
          <View style={styles.legendItem}>
            <MaterialCommunityIcons name="circle" size={10} color={colors.success} />
            <Text style={styles.legendLabel}>Zugeordnet</Text>
          </View>
          <View style={styles.legendItem}>
            <MaterialCommunityIcons name="circle" size={10} color={colors.warning} />
            <Text style={styles.legendLabel}>Aktuell</Text>
          </View>
          <View style={styles.legendItem}>
            <MaterialCommunityIcons name="circle-outline" size={10} color={colors.muted} />
            <Text style={styles.legendLabel}>Offen</Text>
          </View>
        </View>

        <ScrollView
          style={styles.list}
          contentContainerStyle={[
            styles.listContent,
            { paddingBottom: systemBottomInset(insets) + spacing.lg }
          ]}
        >
          {mappingFields.map((field, index) => {
            const status = fieldStatus(field, wizard, currentFieldId);
            const label = resolveFieldDisplayLabel(field, wizard);
            return (
              <Pressable
                key={field.fieldId}
                accessibilityRole="button"
                style={[styles.row, status === 'current' ? styles.rowCurrent : null]}
                onPress={() => {
                  void hapticSelection();
                  onSelectField(index);
                  onClose();
                }}
              >
                <MaterialCommunityIcons
                  name={statusIcon(status)}
                  size={status === 'open' ? 14 : 12}
                  color={statusColor(status)}
                />
                <View style={styles.rowCopy}>
                  <SingleLineText style={styles.rowLabel}>{label}</SingleLineText>
                  <Text style={styles.rowMeta}>Seite {field.page}</Text>
                </View>
                <Text style={styles.rowIndex}>{index + 1}</Text>
              </Pressable>
            );
          })}
        </ScrollView>
      </View>
    </Modal>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg
  },
  header: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panel
  },
  headerBtn: {
    minWidth: 88
  },
  cancel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  legend: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.sm,
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panel
  },
  legendItem: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 6
  },
  legendLabel: {
    ...typography.caption,
    color: colors.muted
  },
  list: {
    flex: 1
  },
  listContent: {
    padding: spacing.pageX,
    gap: spacing.xxs
  },
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.xs,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  rowCurrent: {
    borderColor: colors.warning,
    backgroundColor: 'rgba(196, 140, 40, 0.08)'
  },
  rowCopy: {
    flex: 1,
    gap: 2
  },
  rowLabel: {
    ...typography.body,
    color: colors.ink
  },
  rowMeta: {
    ...typography.caption,
    color: colors.muted
  },
  rowIndex: {
    ...typography.caption,
    color: colors.muted,
    minWidth: 24,
    textAlign: 'right'
  }
});
