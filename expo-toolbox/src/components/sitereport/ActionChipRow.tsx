import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import { hapticSelection } from '../../lib/haptics';

export type ActionChip = {
  key: string;
  label: string;
  icon?: string;
  onPress: () => void;
  disabled?: boolean;
};

type Props = {
  actions: ActionChip[];
};

export function ActionChipRow({ actions }: Props) {
  return (
    <View style={styles.row}>
      {actions.map((action) => (
        <Pressable
          key={action.key}
          accessibilityRole="button"
          disabled={action.disabled}
          onPress={() => {
            void hapticSelection();
            action.onPress();
          }}
          style={({ pressed }) => [
            styles.chip,
            action.disabled ? styles.disabled : null,
            pressed && !action.disabled ? styles.pressed : null
          ]}
        >
          {action.icon ? <Text style={styles.icon}>{action.icon}</Text> : null}
          <Text style={styles.label}>{action.label}</Text>
        </Pressable>
      ))}
    </View>
  );
}

const styles = StyleSheet.create({
  row: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs
  },
  chip: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xs,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.xs,
    borderRadius: spacing.buttonRadius,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border
  },
  pressed: {
    opacity: 0.9,
    backgroundColor: colors.panel
  },
  disabled: {
    opacity: 0.45
  },
  icon: {
    fontSize: 16
  },
  label: {
    ...typography.label,
    color: colors.ink
  }
});
