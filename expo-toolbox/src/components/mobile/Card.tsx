import { ReactNode } from 'react';
import { StyleSheet, View, type ViewStyle } from 'react-native';

import { colors, shadows, spacing } from '../../constants/theme';

type Props = {
  children: ReactNode;
  style?: ViewStyle;
  elevated?: boolean;
  padded?: boolean;
};

export function Card({ children, style, elevated = true, padded = true }: Props) {
  return (
    <View
      style={[
        styles.card,
        elevated ? shadows.card : null,
        padded ? styles.padded : null,
        style
      ]}
    >
      {children}
    </View>
  );
}

const styles = StyleSheet.create({
  card: {
    backgroundColor: colors.panelElevated,
    borderColor: colors.border,
    borderWidth: 1,
    borderRadius: spacing.cardRadius
  },
  padded: {
    padding: spacing.cardPadding
  }
});
