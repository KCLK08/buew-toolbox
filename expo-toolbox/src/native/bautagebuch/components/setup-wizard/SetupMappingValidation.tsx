import { StyleSheet, Text, View } from 'react-native';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import type { MappingTransitionCheck } from '../../lib/setup-mapping';

type Props = {
  check: MappingTransitionCheck;
  onBackToMapping: () => void;
  onContinueAnyway: () => void;
};

export function SetupMappingValidation({ check, onBackToMapping, onContinueAnyway }: Props) {
  return (
    <View style={styles.root}>
      <View style={styles.card}>
        <Text style={styles.warningIcon}>⚠</Text>
        <Text style={styles.title}>Prüfung erforderlich</Text>

        {check.unassignedCount > 0 ? (
          <>
            <Text style={styles.copy}>
              Einige Felder wurden keiner Gruppe zugeordnet. Diese werden später automatisch unter
              „Sonstiges" einsortiert.
            </Text>
            <View style={styles.highlightBox}>
              <Text style={styles.highlightLabel}>Nicht zugeordnet</Text>
              <Text style={styles.highlightValue}>
                {check.unassignedCount} {check.unassignedCount === 1 ? 'Feld' : 'Felder'}
              </Text>
            </View>
          </>
        ) : (
          <>
            {check.issues.map((issue) => (
              <Text key={issue} style={styles.copy}>
                {issue}
              </Text>
            ))}
          </>
        )}

        <View style={styles.actions}>
          <PrimaryButton label="Zurück zur Zuordnung" variant="secondary" onPress={onBackToMapping} />
          <PrimaryButton label="Trotzdem fortfahren" onPress={onContinueAnyway} />
        </View>
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: 'rgba(26, 25, 22, 0.5)',
    justifyContent: 'center',
    padding: spacing.pageX,
    zIndex: 30
  },
  card: {
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panel,
    borderWidth: 1,
    borderColor: colors.border,
    padding: spacing.lg,
    gap: spacing.md
  },
  warningIcon: {
    fontSize: 32,
    textAlign: 'center',
    color: colors.danger
  },
  title: {
    ...typography.title,
    color: colors.ink,
    textAlign: 'center'
  },
  copy: {
    ...typography.body,
    color: colors.muted,
    textAlign: 'center'
  },
  highlightBox: {
    alignItems: 'center',
    gap: spacing.xxs,
    padding: spacing.md,
    borderRadius: 12,
    backgroundColor: colors.badgeBg
  },
  highlightLabel: {
    ...typography.caption,
    color: colors.muted
  },
  highlightValue: {
    ...typography.subtitle,
    color: colors.danger
  },
  actions: {
    gap: spacing.sm,
    marginTop: spacing.xs
  }
});
