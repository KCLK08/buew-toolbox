import { Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { colors, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';

type Props = {
  onEdit: () => void;
  onEditPosition: () => void;
  onDelete: () => void;
  compact?: boolean;
};

export function SetupManualFieldActionMenu({
  onEdit,
  onEditPosition,
  onDelete,
  compact = false
}: Props) {
  return (
    <View style={[styles.root, compact ? styles.rootCompact : null]}>
      <Pressable
        accessibilityRole="button"
        style={styles.action}
        onPress={() => {
          void hapticSelection();
          onEdit();
        }}
      >
        <MaterialCommunityIcons name="pencil-outline" size={18} color={colors.accent} />
        <Text style={styles.actionLabel}>Bearbeiten</Text>
      </Pressable>
      <Pressable
        accessibilityRole="button"
        style={styles.action}
        onPress={() => {
          void hapticSelection();
          onEditPosition();
        }}
      >
        <MaterialCommunityIcons name="vector-square-edit" size={18} color={colors.accent} />
        <Text style={styles.actionLabel}>Position ändern</Text>
      </Pressable>
      <Pressable
        accessibilityRole="button"
        style={styles.action}
        onPress={() => {
          void hapticSelection();
          onDelete();
        }}
      >
        <MaterialCommunityIcons name="trash-can-outline" size={18} color={colors.danger} />
        <Text style={[styles.actionLabel, styles.actionLabelDanger]}>Löschen</Text>
      </Pressable>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs,
    paddingVertical: spacing.xs
  },
  rootCompact: {
    paddingVertical: 0
  },
  action: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs,
    minHeight: 36,
    paddingHorizontal: spacing.sm,
    borderRadius: 999,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated
  },
  actionLabel: {
    ...typography.caption,
    color: colors.accent2,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  actionLabelDanger: {
    color: colors.danger
  }
});
