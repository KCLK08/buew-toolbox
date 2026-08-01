import { Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { colors, spacing, typography } from '../../../../constants/theme';
import { hapticSelection } from '../../../../lib/haptics';

type Props = {
  fieldNumber: number;
  total: number;
  onPrevious?: () => void;
  onNext?: () => void;
  canGoPrevious?: boolean;
  canGoNext?: boolean;
};

export function SetupFieldStepNavigator({
  fieldNumber,
  total,
  onPrevious,
  onNext,
  canGoPrevious = true,
  canGoNext = true
}: Props) {
  if (total <= 0) return null;

  return (
    <View style={styles.root}>
      <Pressable
        accessibilityRole="button"
        accessibilityLabel="Vorheriges Feld"
        style={[styles.navBtn, !canGoPrevious ? styles.navBtnDisabled : null]}
        disabled={!canGoPrevious}
        onPress={() => {
          if (!canGoPrevious) return;
          void hapticSelection();
          onPrevious?.();
        }}
      >
        <MaterialCommunityIcons
          name="chevron-left"
          size={28}
          color={canGoPrevious ? colors.accent : colors.border}
        />
      </Pressable>
      <Text style={styles.counter}>
        Feld {fieldNumber} von {total}
      </Text>
      <Pressable
        accessibilityRole="button"
        accessibilityLabel="Nächstes Feld"
        style={[styles.navBtn, !canGoNext ? styles.navBtnDisabled : null]}
        disabled={!canGoNext}
        onPress={() => {
          if (!canGoNext) return;
          void hapticSelection();
          onNext?.();
        }}
      >
        <MaterialCommunityIcons
          name="chevron-right"
          size={28}
          color={canGoNext ? colors.accent : colors.border}
        />
      </Pressable>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm,
    paddingVertical: spacing.xs
  },
  navBtn: {
    width: 44,
    height: 44,
    alignItems: 'center',
    justifyContent: 'center',
    borderRadius: 999
  },
  navBtnDisabled: {
    opacity: 0.45
  },
  counter: {
    ...typography.subtitle,
    color: colors.ink,
    flex: 1,
    textAlign: 'center'
  }
});
