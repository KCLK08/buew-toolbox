import { ReactNode } from 'react';
import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';

type Props = {
  step: number;
  total: number;
  title: string;
  children: ReactNode;
};

export function WizardStep({ step, total, title, children }: Props) {
  return (
    <View style={styles.wrap}>
      <Text style={styles.progress}>
        Schritt {step}/{total}
      </Text>
      <Text style={styles.title}>{title}</Text>
      <View style={styles.barTrack}>
        <View style={[styles.barFill, { width: `${(step / total) * 100}%` }]} />
      </View>
      {children}
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: {
    gap: spacing.md
  },
  progress: {
    ...typography.label,
    color: colors.accent
  },
  title: {
    ...typography.title,
    color: colors.ink
  },
  barTrack: {
    height: 6,
    borderRadius: 999,
    backgroundColor: colors.border,
    overflow: 'hidden'
  },
  barFill: {
    height: '100%',
    backgroundColor: colors.accent,
    borderRadius: 999
  }
});
