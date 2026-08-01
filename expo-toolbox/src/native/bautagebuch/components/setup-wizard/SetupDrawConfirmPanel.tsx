import { Pressable, StyleSheet, Text, View } from 'react-native';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';

/** Approximate overlay height for PDF scroll reserve (compact panel). */
export const SETUP_DRAFT_CONFIRM_RESERVE_PX = 58;

type Props = {
  rectEditEnabled: boolean;
  onEnableEdit: () => void;
  onConfirm: () => void;
  onCancel: () => void;
  bottomInset?: number;
};

export function SetupDrawConfirmPanel({
  rectEditEnabled,
  onEnableEdit,
  onConfirm,
  onCancel,
  bottomInset = 0
}: Props) {
  return (
    <View style={[styles.root, { paddingBottom: bottomInset + spacing.xxs }]}>
      <View style={styles.row}>
        <Pressable accessibilityRole="button" style={styles.cancelBtn} onPress={onCancel}>
          <Text style={styles.cancelLabel}>Abbrechen</Text>
        </Pressable>
        <View style={styles.actions}>
          <PrimaryButton
            compact
            label={rectEditEnabled ? 'Verschieben…' : 'Verschieben'}
            variant="ghost"
            onPress={onEnableEdit}
          />
          <PrimaryButton compact label="OK" onPress={onConfirm} />
        </View>
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: 'rgba(255, 255, 255, 0.96)',
    paddingHorizontal: spacing.sm,
    paddingTop: spacing.xxs
  },
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs,
    minHeight: 40
  },
  actions: {
    flex: 1,
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'flex-end',
    gap: spacing.xxs
  },
  cancelBtn: {
    minHeight: 36,
    justifyContent: 'center',
    paddingHorizontal: spacing.xxs
  },
  cancelLabel: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 12
  }
});
