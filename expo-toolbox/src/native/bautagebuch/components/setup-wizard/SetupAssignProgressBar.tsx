import { StyleSheet, Text, View } from 'react-native';

import { SingleLineText } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import type { MappingProgress } from '../../lib/setup-mapping';

type Props = {
  progress: MappingProgress;
  fieldNumber?: number;
  fieldLabel?: string | null;
};

export function SetupAssignProgressBar({ progress, fieldNumber, fieldLabel }: Props) {
  const barWidth = `${Math.max(4, progress.percent)}%` as `${number}%`;
  const counter =
    progress.total > 0 && fieldNumber
      ? `Feld ${fieldNumber} von ${progress.total}`
      : progress.total > 0
        ? `${progress.total} erkannte Felder`
        : 'Keine Felder erkannt';

  return (
    <View style={styles.root}>
      <Text style={styles.counter}>{counter}</Text>
      {fieldLabel ? (
        <View style={styles.fieldRow}>
          <Text style={styles.fieldHint}>Erkannt:</Text>
          <SingleLineText style={styles.fieldLabel}>{fieldLabel}</SingleLineText>
        </View>
      ) : null}
      <View style={styles.statsRow}>
        <Text style={styles.stat}>{progress.total} erkannte Felder</Text>
        <Text style={styles.statDot}>·</Text>
        <Text style={[styles.stat, styles.statAssigned]}>{progress.assigned} zugeordnet</Text>
        <Text style={styles.statDot}>·</Text>
        <Text style={styles.stat}>{progress.open} offen</Text>
      </View>
      <View style={styles.track}>
        <View style={[styles.fill, { width: barWidth }]} />
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.xxs,
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.sm,
    backgroundColor: colors.panel,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  counter: {
    ...typography.subtitle,
    color: colors.ink
  },
  fieldRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs,
    marginTop: 2
  },
  fieldHint: {
    ...typography.caption,
    color: colors.muted
  },
  fieldLabel: {
    ...typography.bodyStrong,
    color: colors.accent2,
    flex: 1
  },
  statsRow: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    alignItems: 'center',
    gap: 4,
    marginTop: spacing.xxs
  },
  stat: {
    ...typography.caption,
    color: colors.muted
  },
  statAssigned: {
    color: colors.success
  },
  statDot: {
    ...typography.caption,
    color: colors.border
  },
  track: {
    height: 8,
    borderRadius: 999,
    backgroundColor: colors.border,
    overflow: 'hidden',
    marginTop: spacing.xxs
  },
  fill: {
    height: '100%',
    borderRadius: 999,
    backgroundColor: colors.accent
  }
});
