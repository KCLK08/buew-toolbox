import { Modal, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import {
  SETUP_FIELD_TYPE_OPTIONS,
  setupFieldTypeLabel
} from '../../lib/setup-field-settings';
import type { SetupFieldType } from '../../types';

type Props = {
  visible: boolean;
  value: SetupFieldType;
  onClose: () => void;
  onSelect: (type: SetupFieldType) => void;
};

export function SetupFieldTypePickerSheet({ visible, value, onClose, onSelect }: Props) {
  const insets = useSafeAreaInsets();

  return (
    <Modal visible={visible} animationType="slide" transparent onRequestClose={onClose}>
      <Pressable style={styles.backdrop} onPress={onClose} />
      <View style={[styles.sheet, { paddingBottom: systemBottomInset(insets) + spacing.sm }]}>
        <View style={styles.handle} />
        <Text style={styles.title}>Feldtyp</Text>
        <Text style={styles.subtitle}>Aktuell: {setupFieldTypeLabel(value)}</Text>
        <ScrollView style={styles.list} showsVerticalScrollIndicator={false}>
          {SETUP_FIELD_TYPE_OPTIONS.map((option) => {
            const active = option.value === value;
            return (
              <Pressable
                key={option.value}
                accessibilityRole="button"
                style={[styles.row, active ? styles.rowActive : null]}
                onPress={() => {
                  void hapticSelection();
                  onSelect(option.value);
                  onClose();
                }}
              >
                <MaterialCommunityIcons
                  name={active ? 'radiobox-marked' : 'radiobox-blank'}
                  size={20}
                  color={active ? colors.accent : colors.muted}
                />
                <Text style={active ? styles.rowLabelActive : styles.rowLabel}>{option.label}</Text>
              </Pressable>
            );
          })}
        </ScrollView>
      </View>
    </Modal>
  );
}

const styles = StyleSheet.create({
  backdrop: {
    flex: 1,
    backgroundColor: 'rgba(26, 25, 22, 0.35)'
  },
  sheet: {
    backgroundColor: colors.panel,
    borderTopLeftRadius: 18,
    borderTopRightRadius: 18,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    maxHeight: '70%'
  },
  handle: {
    alignSelf: 'center',
    width: 40,
    height: 4,
    borderRadius: 999,
    backgroundColor: colors.border,
    marginBottom: spacing.sm
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  subtitle: {
    ...typography.caption,
    color: colors.muted,
    marginBottom: spacing.sm
  },
  list: {
    flexGrow: 0
  },
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.sm,
    borderRadius: 12,
    marginBottom: spacing.xxs
  },
  rowActive: {
    backgroundColor: colors.badgeBg
  },
  rowLabel: {
    ...typography.body,
    color: colors.ink
  },
  rowLabelActive: {
    ...typography.bodyStrong,
    color: colors.accent2
  }
});
