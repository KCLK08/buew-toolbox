import { Pressable, StyleSheet, Text, View } from 'react-native';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';

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
    <View style={[styles.root, { paddingBottom: bottomInset + spacing.sm }]}>
      <View style={styles.card}>
        <Text style={styles.title}>Feldposition prüfen</Text>
        <Text style={styles.copy}>Ist die Markierung korrekt?</Text>
        <View style={styles.actions}>
          <Pressable accessibilityRole="button" style={styles.cancelBtn} onPress={onCancel}>
            <Text style={styles.cancelLabel}>Abbrechen</Text>
          </Pressable>
          <PrimaryButton
            compact
            label={rectEditEnabled ? 'Verschieben…' : 'Verschieben'}
            variant="ghost"
            onPress={onEnableEdit}
          />
          <PrimaryButton compact label="Bestätigen" onPress={onConfirm} />
        </View>
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm
  },
  card: {
    gap: spacing.xs
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  copy: {
    ...typography.body,
    color: colors.muted
  },
  actions: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'flex-end',
    flexWrap: 'wrap',
    gap: spacing.xs,
    marginTop: spacing.xxs
  },
  cancelBtn: {
    minHeight: 36,
    justifyContent: 'center',
    paddingHorizontal: spacing.xs,
    marginRight: 'auto'
  },
  cancelLabel: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  }
});
