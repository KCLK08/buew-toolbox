import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import { StatusBadge } from '../mobile';

type Props = {
  title: string;
  projectName: string;
  date: string;
  statusLabel?: string;
  statusTone?: 'warning' | 'info' | 'success' | 'neutral';
};

export function ProtocolHero({ title, projectName, date, statusLabel, statusTone = 'neutral' }: Props) {
  return (
    <View style={styles.hero}>
      <Text style={styles.title}>{title}</Text>
      <Text style={styles.project}>{projectName}</Text>
      <View style={styles.meta}>
        <Text style={styles.date}>{date}</Text>
        {statusLabel ? <StatusBadge label={statusLabel} tone={statusTone} /> : null}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  hero: {
    gap: spacing.xs,
    marginBottom: spacing.md
  },
  title: {
    ...typography.title,
    color: colors.ink
  },
  project: {
    ...typography.body,
    color: colors.muted
  },
  meta: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    marginTop: spacing.xs
  },
  date: {
    ...typography.label,
    color: colors.accent2
  }
});
