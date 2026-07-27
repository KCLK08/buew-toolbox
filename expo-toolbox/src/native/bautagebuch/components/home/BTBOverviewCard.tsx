import { Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { Card } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import type { BtbHomeStats } from '../../lib/home-utils';

type Props = {
  stats: BtbHomeStats;
  onPress: () => void;
};

function StatBlock({ label, value }: { label: string; value: number }) {
  return (
    <View style={styles.statBlock}>
      <Text style={styles.statLabel}>{label}</Text>
      <Text style={styles.statValue}>{value}</Text>
    </View>
  );
}

export function BTBOverviewCard({ stats, onPress }: Props) {
  return (
    <Pressable
      accessibilityRole="button"
      accessibilityLabel="Übersicht aller Bautagebücher öffnen"
      onPress={onPress}
    >
      <Card style={styles.card}>
        <View style={styles.header}>
          <Text style={styles.title}>Übersicht</Text>
          <MaterialCommunityIcons name="chevron-right" size={22} color={colors.muted} />
        </View>
        <View style={styles.statsRow}>
          <StatBlock label="Gesamt" value={stats.total} />
          <StatBlock label="Entwürfe" value={stats.drafts} />
          <StatBlock label="Abgeschlossen" value={stats.completed} />
        </View>
      </Card>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  card: {
    gap: spacing.md,
    backgroundColor: colors.panelElevated
  },
  header: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between'
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  statsRow: {
    flexDirection: 'row',
    gap: spacing.sm
  },
  statBlock: {
    flex: 1,
    gap: spacing.xxs,
    minHeight: spacing.touchMin,
    justifyContent: 'center',
    paddingVertical: spacing.xs,
    paddingHorizontal: spacing.sm,
    borderRadius: 12,
    backgroundColor: colors.bg,
    borderWidth: 1,
    borderColor: colors.border
  },
  statLabel: {
    ...typography.caption,
    color: colors.muted
  },
  statValue: {
    ...typography.title,
    color: colors.ink,
    fontSize: 24,
    lineHeight: 28
  }
});
