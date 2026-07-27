import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../../../constants/theme';
import type { MappingProgress } from '../../lib/setup-mapping';

type Props = {
  progress: MappingProgress;
  fieldNumber?: number;
  fieldName?: string;
  labelCandidate?: string;
};

export function SetupProgressHeader({
  progress,
  fieldNumber,
  fieldName,
  labelCandidate
}: Props) {
  const barWidth = `${Math.max(4, progress.percent)}%` as `${number}%`;
  const counter =
    progress.total > 0 && fieldNumber
      ? `Feld ${fieldNumber} von ${progress.total}`
      : progress.total > 0
        ? `Feld ${Math.min(progress.current, progress.total)} von ${progress.total}`
        : 'Keine Felder';

  return (
    <View style={styles.root}>
      <Text style={styles.counter}>{counter}</Text>
      <View style={styles.track}>
        <View style={[styles.fill, { width: barWidth }]} />
      </View>
      <Text style={styles.percent}>{progress.percent}%</Text>

      {fieldName ? (
        <View style={styles.fieldInfo}>
          <View style={styles.fieldRow}>
            <Text style={styles.fieldLabel}>Erkannter Name:</Text>
            <Text style={styles.fieldValue} numberOfLines={2}>
              {fieldName}
            </Text>
          </View>
          {labelCandidate && labelCandidate !== fieldName ? (
            <View style={styles.fieldRow}>
              <Text style={styles.fieldLabel}>Vorschlag:</Text>
              <Text style={styles.fieldValueStrong} numberOfLines={2}>
                {labelCandidate}
              </Text>
            </View>
          ) : null}
        </View>
      ) : null}
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.xs,
    paddingHorizontal: spacing.pageX,
    paddingTop: spacing.sm,
    paddingBottom: spacing.sm,
    backgroundColor: colors.panel,
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  counter: {
    ...typography.subtitle,
    color: colors.ink
  },
  track: {
    height: 12,
    borderRadius: 999,
    backgroundColor: colors.border,
    overflow: 'hidden'
  },
  fill: {
    height: '100%',
    borderRadius: 999,
    backgroundColor: colors.accent
  },
  percent: {
    ...typography.caption,
    color: colors.muted
  },
  fieldInfo: {
    gap: spacing.xs,
    marginTop: spacing.xxs,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border
  },
  fieldRow: {
    gap: 2
  },
  fieldLabel: {
    ...typography.caption,
    color: colors.muted
  },
  fieldValue: {
    ...typography.body,
    color: colors.ink
  },
  fieldValueStrong: {
    ...typography.bodyStrong,
    color: colors.accent2
  }
});
