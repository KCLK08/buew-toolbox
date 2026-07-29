import { StyleSheet, Text, View } from 'react-native';

import { SingleLineText } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import type { FieldSettingsProgress } from '../../lib/setup-field-settings';

type Props = {
  progress: FieldSettingsProgress;
  fieldNumber?: number;
  fieldLabel?: string | null;
  detectedTypeLabel?: string | null;
};

export function SetupFieldsProgressBar({
  progress,
  fieldNumber,
  fieldLabel,
  detectedTypeLabel
}: Props) {
  const barWidth = `${Math.max(4, progress.percent)}%` as `${number}%`;
  const counter =
    progress.total > 0 && fieldNumber
      ? `Feld ${fieldNumber} von ${progress.total}`
      : `${progress.total} Felder`;

  return (
    <View style={styles.root}>
      <Text style={styles.counter}>{counter}</Text>
      {fieldLabel ? (
        <SingleLineText style={styles.fieldLabel}>{`„${fieldLabel}"`}</SingleLineText>
      ) : null}
      {detectedTypeLabel ? (
        <View style={styles.detectedRow}>
          <Text style={styles.detectedHint}>Erkannt:</Text>
          <Text style={styles.detectedValue}>{detectedTypeLabel}</Text>
        </View>
      ) : null}
      <View style={styles.statsRow}>
        <Text style={styles.stat}>{progress.total} Felder</Text>
        <Text style={styles.statDot}>·</Text>
        <Text style={[styles.stat, styles.statConfigured]}>{progress.configured} konfiguriert</Text>
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
  fieldLabel: {
    ...typography.bodyStrong,
    color: colors.accent2
  },
  detectedRow: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs
  },
  detectedHint: {
    ...typography.caption,
    color: colors.muted
  },
  detectedValue: {
    ...typography.caption,
    color: colors.ink,
    fontFamily: 'SpaceGrotesk_600SemiBold'
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
  statConfigured: {
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
