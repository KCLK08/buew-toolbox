import { ScrollView, StyleSheet, Text, View } from 'react-native';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import type { MappingCompletionSummary } from '../../lib/setup-mapping';

type Props = {
  templateName: string;
  summary: MappingCompletionSummary;
  onConfigureFields: () => void;
  onFinishLater: () => void;
};

export function SetupMappingCompletion({
  templateName,
  summary,
  onConfigureFields,
  onFinishLater
}: Props) {
  return (
    <ScrollView
      style={styles.root}
      contentContainerStyle={styles.content}
      keyboardShouldPersistTaps="handled"
    >
      <View style={styles.hero}>
        <Text style={styles.checkmark}>✓</Text>
        <Text style={styles.title}>Feldzuordnung abgeschlossen</Text>
        <Text style={styles.subtitle}>Die PDF-Vorlage wurde erfolgreich analysiert.</Text>
      </View>

      <View style={styles.statsCard}>
        <View style={styles.statRow}>
          <Text style={styles.statLabel}>Vorlage</Text>
          <Text style={styles.statValue}>{templateName}</Text>
        </View>
        <View style={styles.statRow}>
          <Text style={styles.statLabel}>Erkannte Felder</Text>
          <Text style={styles.statValue}>{summary.totalFields}</Text>
        </View>
        <View style={styles.statRow}>
          <Text style={styles.statLabel}>Gruppen</Text>
          <Text style={styles.statValue}>{summary.groupCount}</Text>
        </View>
        <View style={styles.statRow}>
          <Text style={styles.statLabel}>Nicht zugeordnet</Text>
          <Text style={[styles.statValue, summary.unassignedCount > 0 ? styles.warn : null]}>
            {summary.unassignedCount}
          </Text>
        </View>
      </View>

      <Text style={styles.sectionHeading}>Gruppenübersicht</Text>
      <View style={styles.groupList}>
        {summary.groups.map((group) => (
          <View key={group.sectionId} style={styles.groupRow}>
            <Text style={styles.groupCheck}>✓</Text>
            <View style={styles.groupCopy}>
              <Text style={styles.groupLabel}>{group.label}</Text>
              <Text style={styles.groupMeta}>
                {group.fieldCount} {group.fieldCount === 1 ? 'Feld' : 'Felder'}
              </Text>
            </View>
          </View>
        ))}
      </View>

      <View style={styles.actions}>
        <PrimaryButton label="Felder konfigurieren" onPress={onConfigureFields} />
        <PrimaryButton label="Später fortsetzen" variant="ghost" onPress={onFinishLater} />
      </View>
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  root: { flex: 1 },
  content: {
    padding: spacing.pageX,
    paddingBottom: spacing.xxl,
    gap: spacing.md
  },
  hero: {
    alignItems: 'center',
    gap: spacing.xs,
    paddingVertical: spacing.md
  },
  checkmark: {
    fontSize: 36,
    color: colors.success,
    fontFamily: 'SpaceGrotesk_700Bold'
  },
  title: {
    ...typography.title,
    color: colors.ink,
    textAlign: 'center'
  },
  subtitle: {
    ...typography.body,
    color: colors.muted,
    textAlign: 'center'
  },
  statsCard: {
    borderRadius: spacing.cardRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel,
    padding: spacing.md,
    gap: spacing.sm
  },
  statRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: spacing.sm
  },
  statLabel: {
    ...typography.caption,
    color: colors.muted
  },
  statValue: {
    ...typography.bodyStrong,
    color: colors.ink,
    flexShrink: 1,
    textAlign: 'right'
  },
  warn: {
    color: colors.danger
  },
  sectionHeading: {
    ...typography.subtitle,
    color: colors.ink
  },
  groupList: {
    gap: spacing.sm
  },
  groupRow: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    gap: spacing.sm,
    padding: spacing.sm,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  groupCheck: {
    ...typography.bodyStrong,
    color: colors.accent,
    marginTop: 2
  },
  groupCopy: {
    flex: 1,
    gap: 2
  },
  groupLabel: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  groupMeta: {
    ...typography.caption,
    color: colors.muted
  },
  actions: {
    gap: spacing.sm,
    marginTop: spacing.sm
  }
});
