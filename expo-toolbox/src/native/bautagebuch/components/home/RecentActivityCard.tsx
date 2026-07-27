import { Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { Card } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { displayRunTitle, formatActivityTime, runActivityMeta } from '../../lib/home-utils';
import type { BautagebuchRun } from '../../types';

type Props = {
  runs: BautagebuchRun[];
  onOpenRun: (runId: string) => void;
};

export function RecentActivityCard({ runs, onOpenRun }: Props) {
  return (
    <Card style={styles.card}>
      <Text style={styles.title}>Letzte Aktivität</Text>

      {runs.length === 0 ? (
        <View style={styles.empty}>
          <Text style={styles.emptyText}>Noch keine Bautagebücher vorhanden</Text>
        </View>
      ) : (
        <View style={styles.list}>
          {runs.map((run) => {
            const meta = runActivityMeta(run);
            return (
              <Pressable
                key={run.runId}
                accessibilityRole="button"
                style={styles.row}
                onPress={() => onOpenRun(run.runId)}
              >
                <View style={styles.iconWrap}>
                  <MaterialCommunityIcons name="notebook-outline" size={20} color={colors.accent} />
                </View>
                <View style={styles.rowCopy}>
                  <Text style={styles.rowTitle} numberOfLines={2}>
                    {displayRunTitle(run.title)}
                  </Text>
                  <Text style={styles.rowMeta}>{formatActivityTime(run.updatedAt)}</Text>
                  {meta ? <Text style={styles.rowExtra}>{meta}</Text> : null}
                </View>
                <MaterialCommunityIcons name="chevron-right" size={20} color={colors.muted} />
              </Pressable>
            );
          })}
        </View>
      )}
    </Card>
  );
}

const styles = StyleSheet.create({
  card: {
    gap: spacing.sm,
    backgroundColor: colors.panelElevated
  },
  title: {
    ...typography.subtitle,
    color: colors.ink
  },
  empty: {
    minHeight: spacing.touchMin + spacing.sm,
    alignItems: 'center',
    justifyContent: 'center',
    paddingVertical: spacing.md
  },
  emptyText: {
    ...typography.body,
    color: colors.muted,
    textAlign: 'center'
  },
  list: {
    gap: spacing.xs
  },
  row: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: spacing.touchMin + 8,
    paddingVertical: spacing.sm,
    paddingHorizontal: spacing.xs,
    borderRadius: 12,
    backgroundColor: colors.bg,
    borderWidth: 1,
    borderColor: colors.border
  },
  iconWrap: {
    width: 36,
    height: 36,
    borderRadius: 10,
    backgroundColor: colors.badgeBg,
    alignItems: 'center',
    justifyContent: 'center'
  },
  rowCopy: {
    flex: 1,
    gap: 2
  },
  rowTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  rowMeta: {
    ...typography.caption,
    color: colors.muted
  },
  rowExtra: {
    ...typography.caption,
    color: colors.accent2
  }
});
