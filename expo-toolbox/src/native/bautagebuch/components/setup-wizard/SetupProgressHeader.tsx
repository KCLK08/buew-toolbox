import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../../../constants/theme';
import type { MappingProgress } from '../../lib/setup-mapping';

type Props = {
  progress: MappingProgress;
  title?: string;
};

export function SetupProgressHeader({ progress, title }: Props) {
  const barWidth = `${Math.max(4, progress.percent)}%` as `${number}%`;

  return (
    <View style={styles.root}>
      {title ? <Text style={styles.title}>{title}</Text> : null}
      <Text style={styles.counter}>
        Feld {Math.min(progress.current, progress.total)} von {progress.total}
      </Text>
      <View style={styles.track}>
        <View style={[styles.fill, { width: barWidth }]} />
      </View>
      <View style={styles.metaRow}>
        <Text style={styles.meta}>{progress.percent}%</Text>
        <Text style={styles.meta}>{progress.remaining} verbleibend</Text>
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.xs,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    paddingBottom: spacing.xs
  },
  title: {
    ...typography.caption,
    color: colors.muted
  },
  counter: {
    ...typography.subtitle,
    color: colors.ink
  },
  track: {
    height: 10,
    borderRadius: 999,
    backgroundColor: colors.border,
    overflow: 'hidden'
  },
  fill: {
    height: '100%',
    borderRadius: 999,
    backgroundColor: colors.accent
  },
  metaRow: {
    flexDirection: 'row',
    justifyContent: 'space-between'
  },
  meta: {
    ...typography.caption,
    color: colors.muted
  }
});
