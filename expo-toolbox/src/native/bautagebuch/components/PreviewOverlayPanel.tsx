import { ReactNode } from 'react';
import { Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../../constants/theme';

type Props = {
  title?: string;
  onClose?: () => void;
  children: ReactNode;
};

export function PreviewOverlayPanel({ title = 'Live-Vorschau', onClose, children }: Props) {
  return (
    <View style={styles.overlay}>
      <View style={styles.panel}>
        <View style={styles.header}>
          <Text style={styles.title}>{title}</Text>
          {onClose ? (
            <Pressable
              accessibilityRole="button"
              accessibilityLabel="Vorschau schließen"
              hitSlop={8}
              onPress={onClose}
            >
              <Text style={styles.close}>Schließen</Text>
            </Pressable>
          ) : null}
        </View>
        <ScrollView
          style={styles.body}
          contentContainerStyle={styles.bodyContent}
          keyboardShouldPersistTaps="handled"
          showsVerticalScrollIndicator={false}
        >
          {children}
        </ScrollView>
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  overlay: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: 'rgba(26, 25, 22, 0.45)',
    padding: spacing.sm,
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
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.xs,
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panelElevated
  },
  title: {
    ...typography.label,
    color: colors.ink
  },
  close: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  body: {
    flex: 1
  },
  bodyContent: {
    padding: spacing.sm,
    gap: spacing.sm
  }
});
