import { ReactNode } from 'react';
import { Modal, Pressable, StyleSheet, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../constants/theme';
import { systemBottomInset } from '../../navigation/systemInsets';
import { PrimaryButton } from '../mobile';

export type ConfirmAction = {
  label: string;
  onPress: () => void;
  variant?: 'primary' | 'secondary' | 'ghost' | 'danger';
};

type Props = {
  visible: boolean;
  title: string;
  message?: string;
  actions: ConfirmAction[];
  onClose: () => void;
};

export function ConfirmModal({ visible, title, message, actions, onClose }: Props) {
  return (
    <Modal visible={visible} transparent animationType="fade" onRequestClose={onClose}>
      <Pressable style={styles.backdrop} onPress={onClose}>
        <Pressable style={styles.sheet} onPress={(e) => e.stopPropagation()}>
          <Text style={styles.title}>{title}</Text>
          {message ? <Text style={styles.message}>{message}</Text> : null}
          <View style={styles.actions}>
            {actions.map((action) => (
              <PrimaryButton
                key={action.label}
                label={action.label}
                variant={action.variant ?? 'secondary'}
                onPress={action.onPress}
              />
            ))}
          </View>
        </Pressable>
      </Pressable>
    </Modal>
  );
}

type BottomSheetProps = {
  visible: boolean;
  title: string;
  children: ReactNode;
  onClose: () => void;
};

export function BottomSheet({ visible, title, children, onClose }: BottomSheetProps) {
  const insets = useSafeAreaInsets();

  return (
    <Modal visible={visible} transparent animationType="slide" onRequestClose={onClose}>
      <Pressable style={styles.backdrop} onPress={onClose}>
        <Pressable
          style={[styles.bottomSheet, { paddingBottom: systemBottomInset(insets) + spacing.lg }]}
          onPress={(e) => e.stopPropagation()}
        >
          <View style={styles.handle} />
          <Text style={styles.title}>{title}</Text>
          {children}
        </Pressable>
      </Pressable>
    </Modal>
  );
}

const styles = StyleSheet.create({
  backdrop: {
    flex: 1,
    backgroundColor: colors.overlay,
    justifyContent: 'center',
    padding: spacing.pageX
  },
  sheet: {
    backgroundColor: colors.panel,
    borderRadius: spacing.cardRadius,
    padding: spacing.lg,
    gap: spacing.sm,
    borderWidth: 1,
    borderColor: colors.border
  },
  bottomSheet: {
    marginTop: 'auto',
    backgroundColor: colors.panel,
    borderTopLeftRadius: spacing.cardRadius + 4,
    borderTopRightRadius: spacing.cardRadius + 4,
    padding: spacing.lg,
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
  message: {
    ...typography.body,
    color: colors.muted
  },
  actions: {
    gap: spacing.xs,
    marginTop: spacing.xs
  }
});
