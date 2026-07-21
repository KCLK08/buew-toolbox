import { ReactNode } from 'react';
import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';

type Props = {
  title: string;
  action?: ReactNode;
  children?: ReactNode;
};

export function Section({ title, action, children }: Props) {
  return (
    <View style={styles.section}>
      <View style={styles.header}>
        <Text style={styles.title}>{title}</Text>
        {action}
      </View>
      {children}
    </View>
  );
}

const styles = StyleSheet.create({
  section: {
    gap: spacing.sm,
    marginBottom: spacing.sectionGap
  },
  header: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingHorizontal: spacing.xxs
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  }
});
