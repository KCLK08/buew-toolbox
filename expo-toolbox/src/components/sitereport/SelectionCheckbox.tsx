import { Pressable, StyleSheet, Text } from 'react-native';

import { colors } from '../../constants/theme';
import { hapticSelection } from '../../lib/haptics';

type Props = {
  selected?: boolean;
  onToggle: () => void;
};

export function SelectionCheckbox({ selected, onToggle }: Props) {
  return (
    <Pressable
      accessibilityRole="checkbox"
      accessibilityState={{ checked: selected }}
      hitSlop={8}
      onPress={() => {
        void hapticSelection();
        onToggle();
      }}
      style={styles.wrap}
    >
      <Text style={[styles.box, selected ? styles.on : null]}>{selected ? '✓' : ''}</Text>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  wrap: {
    paddingTop: 2
  },
  box: {
    width: 26,
    height: 26,
    borderRadius: 8,
    borderWidth: 2,
    borderColor: colors.borderStrong,
    backgroundColor: colors.panelElevated,
    textAlign: 'center',
    lineHeight: 22,
    fontSize: 14,
    fontWeight: '700',
    color: colors.white,
    overflow: 'hidden'
  },
  on: {
    borderColor: colors.accent,
    backgroundColor: colors.accent
  }
});
