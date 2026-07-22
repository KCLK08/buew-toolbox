import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../../constants/theme';
import {
  formatRunCount,
  type WeekRunGroup,
  type YearRunGroup
} from '../lib/group-runs-by-calendar';
import type { BautagebuchRun } from '../types';
import { BautagebuchRunCard } from './BautagebuchRunCard';

type Props = {
  tree: { multiYear: boolean; years: YearRunGroup[] };
  expandedYears: Set<number>;
  expandedWeeks: Set<string>;
  selectionMode: boolean;
  selectedRunIds: string[];
  onToggleYear: (year: number) => void;
  onToggleWeek: (weekKey: string) => void;
  onOpenRun: (runId: string) => void;
  onToggleSelect: (runId: string) => void;
  onRename: (run: BautagebuchRun) => void;
  onDelete: (runId: string) => void;
};

function WeekBlock({
  week,
  expanded,
  selectionMode,
  selectedRunIds,
  onToggleWeek,
  onOpenRun,
  onToggleSelect,
  onRename,
  onDelete
}: {
  week: WeekRunGroup;
  expanded: boolean;
  selectionMode: boolean;
  selectedRunIds: string[];
  onToggleWeek: (weekKey: string) => void;
  onOpenRun: (runId: string) => void;
  onToggleSelect: (runId: string) => void;
  onRename: (run: BautagebuchRun) => void;
  onDelete: (runId: string) => void;
}) {
  return (
    <View style={styles.weekGroup}>
      <Pressable style={styles.weekHeader} onPress={() => onToggleWeek(week.weekKey)}>
        <View style={styles.weekHeaderMain}>
          <Text style={styles.weekTitle}>{week.weekLabel}</Text>
          <Text style={styles.weekRange}>{week.dateRangeLabel}</Text>
        </View>
        <Text style={styles.weekMeta}>
          {expanded ? '▾' : '▸'} {formatRunCount(week.runs.length)}
        </Text>
      </Pressable>
      {expanded
        ? week.runs.map((run) => (
            <BautagebuchRunCard
              key={run.runId}
              run={run}
              selectionMode={selectionMode}
              selected={selectedRunIds.includes(run.runId)}
              onPress={() => onOpenRun(run.runId)}
              onToggleSelect={() => onToggleSelect(run.runId)}
              onRename={() => onRename(run)}
              onDelete={() => onDelete(run.runId)}
            />
          ))
        : null}
    </View>
  );
}

export function BautagebuchRunList({
  tree,
  expandedYears,
  expandedWeeks,
  selectionMode,
  selectedRunIds,
  onToggleYear,
  onToggleWeek,
  onOpenRun,
  onToggleSelect,
  onRename,
  onDelete
}: Props) {
  if (tree.multiYear) {
    return (
      <View style={styles.root}>
        {tree.years.map((yearGroup) => {
          const yearExpanded = expandedYears.has(yearGroup.year);
          return (
            <View key={yearGroup.year} style={styles.yearGroup}>
              <Pressable style={styles.yearHeader} onPress={() => onToggleYear(yearGroup.year)}>
                <Text style={styles.yearTitle}>{yearGroup.year > 0 ? String(yearGroup.year) : 'Ohne Jahr'}</Text>
                <Text style={styles.yearMeta}>
                  {yearExpanded ? '▾' : '▸'} {formatRunCount(yearGroup.runCount)}
                </Text>
              </Pressable>
              {yearExpanded
                ? yearGroup.weeks.map((week) => (
                    <WeekBlock
                      key={week.weekKey}
                      week={week}
                      expanded={expandedWeeks.has(week.weekKey)}
                      selectionMode={selectionMode}
                      selectedRunIds={selectedRunIds}
                      onToggleWeek={onToggleWeek}
                      onOpenRun={onOpenRun}
                      onToggleSelect={onToggleSelect}
                      onRename={onRename}
                      onDelete={onDelete}
                    />
                  ))
                : null}
            </View>
          );
        })}
      </View>
    );
  }

  const weeks = tree.years[0]?.weeks || tree.years.flatMap((year) => year.weeks);

  return (
    <View style={styles.root}>
      {weeks.map((week) => (
        <WeekBlock
          key={week.weekKey}
          week={week}
          expanded={expandedWeeks.has(week.weekKey)}
          selectionMode={selectionMode}
          selectedRunIds={selectedRunIds}
          onToggleWeek={onToggleWeek}
          onOpenRun={onOpenRun}
          onToggleSelect={onToggleSelect}
          onRename={onRename}
          onDelete={onDelete}
        />
      ))}
    </View>
  );
}

const styles = StyleSheet.create({
  root: { gap: spacing.md },
  yearGroup: { gap: spacing.sm },
  yearHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.xxs,
    paddingTop: spacing.xs
  },
  yearTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  yearMeta: {
    ...typography.caption,
    color: colors.muted
  },
  weekGroup: {
    gap: spacing.sm,
    marginLeft: spacing.xs
  },
  weekHeader: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    gap: spacing.sm,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.xxs,
    paddingVertical: spacing.xxs
  },
  weekHeaderMain: {
    flex: 1,
    gap: 2
  },
  weekTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  weekRange: {
    ...typography.caption,
    color: colors.muted,
    fontFamily: 'SpaceGrotesk_400Regular'
  },
  weekMeta: {
    ...typography.caption,
    color: colors.muted,
    paddingTop: 2
  }
});
