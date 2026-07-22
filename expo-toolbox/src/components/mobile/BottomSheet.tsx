import { ReactNode } from 'react';
import { Modal, Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import { PrimaryButton } from './PrimaryButton';

type BottomSheetProps = {
  visible: boolean;
  title: string;
  subtitle?: string;
  children: ReactNode;
  onClose: () => void;
};

export function BottomSheet({ visible, title, subtitle, children, onClose }: BottomSheetProps) {
  return (
    <Modal visible={visible} transparent animationType="slide" onRequestClose={onClose}>
      <Pressable style={styles.backdrop} onPress={onClose}>
        <Pressable style={styles.sheet} onPress={(event) => event.stopPropagation()}>
          <View style={styles.handle} />
          <Text style={styles.title}>{title}</Text>
          {subtitle ? <Text style={styles.subtitle}>{subtitle}</Text> : null}
          {children}
        </Pressable>
      </Pressable>
    </Modal>
  );
}

type OptionProps = {
  label: string;
  description?: string;
  selected: boolean;
  disabled?: boolean;
  onPress: () => void;
};

export function BottomSheetOption({ label, description, selected, disabled = false, onPress }: OptionProps) {
  return (
    <Pressable
      accessibilityRole="radio"
      accessibilityState={{ selected, disabled }}
      disabled={disabled}
      onPress={onPress}
      style={[
        styles.option,
        selected ? styles.optionSelected : null,
        disabled ? styles.optionDisabled : null
      ]}
    >
      <View style={[styles.radio, selected ? styles.radioSelected : null]}>
        {selected ? <View style={styles.radioDot} /> : null}
      </View>
      <View style={styles.optionText}>
        <Text style={styles.optionLabel}>{label}</Text>
        {description ? <Text style={styles.optionDescription}>{description}</Text> : null}
      </View>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  backdrop: {
    flex: 1,
    backgroundColor: colors.overlay,
    justifyContent: 'flex-end'
  },
  sheet: {
    backgroundColor: colors.panel,
    borderTopLeftRadius: spacing.cardRadius + 4,
    borderTopRightRadius: spacing.cardRadius + 4,
    paddingHorizontal: spacing.lg,
    paddingTop: spacing.sm,
    paddingBottom: spacing.xxl,
    gap: spacing.sm,
    borderWidth: 1,
    borderColor: colors.border
  },
  handle: {
    alignSelf: 'center',
    width: 40,
    height: 4,
    borderRadius: 2,
    backgroundColor: colors.borderStrong,
    marginBottom: spacing.xs
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  subtitle: {
    ...typography.body,
    color: colors.muted
  },
  option: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.sm,
    padding: spacing.md,
    borderRadius: spacing.inputRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    minHeight: spacing.touchMin + 8
  },
  optionSelected: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  optionDisabled: {
    opacity: 0.45
  },
  radio: {
    width: 22,
    height: 22,
    borderRadius: 11,
    borderWidth: 2,
    borderColor: colors.borderStrong,
    alignItems: 'center',
    justifyContent: 'center',
    marginTop: 2
  },
  radioSelected: {
    borderColor: colors.accent
  },
  radioDot: {
    width: 10,
    height: 10,
    borderRadius: 5,
    backgroundColor: colors.accent
  },
  optionText: {
    flex: 1,
    gap: 2
  },
  optionLabel: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  optionDescription: {
    ...typography.caption,
    color: colors.muted
  }
});
