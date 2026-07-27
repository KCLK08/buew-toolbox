import { ReactNode } from 'react';
import { Pressable, StyleSheet, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../constants/theme';

type Props = {
  title?: string;
  onClose?: () => void;
  children: ReactNode;
};

export function PreviewOverlayPanel({ title = 'Live-Vorschau', onClose, children }: Props) {
  const insets = useSafeAreaInsets();
  const topInset = Math.max(insets.top, spacing.xs);

  return (
    <View style={styles.overlay}>
      <View
        style={[
          styles.panel,
          {
            marginTop: topInset,
            marginHorizontal: spacing.xs
          }
        ]}
      >
        <View style={styles.header}>
          <Text style={styles.title}>{title}</Text>
          {onClose ? (
            <Pressable
              accessibilityRole="button"
              accessibilityLabel="Vorschau schließen"
              hitSlop={12}
              style={styles.closeBtn}
              onPress={onClose}
            >
              <Text style={styles.close}>Schließen</Text>
            </Pressable>
          ) : null}
        </View>
        <View style={styles.body}>{children}</View>
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  overlay: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: 'rgba(26, 25, 22, 0.42)',
    zIndex: 20
  },
  panel: {
    flex: 1,
    borderRadius: 14,
    overflow: 'hidden',
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border
  },
  header: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm,
    paddingHorizontal: spacing.md,
    minHeight: spacing.touchMin,
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panelElevated
  },
  title: {
    ...typography.label,
    color: colors.ink,
    flex: 1
  },
  closeBtn: {
    minHeight: spacing.touchMin,
    justifyContent: 'center',
    paddingHorizontal: spacing.xs
  },
  close: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  body: {
    flex: 1,
    minHeight: 0
  }
});
