import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../../constants/theme';
import {
  formatRunCount,
  projectGroupKey,
  type ProjectRunGroup,
  type WeekRunGroup,
  type YearRunGroup
} from '../lib/group-runs-by-calendar';
import type { BautagebuchRun } from '../types';
import { BautagebuchRunCard } from './BautagebuchRunCard';

type Props = {
  tree: { multiYear: boolean; years: YearRunGroup[] };
  expandedYears: Set<number>;
  expandedWeeks: Set<string>;
  expandedProjects: Set<string>;
  selectionMode: boolean;
  selectedRunIds: string[];
  onToggleYear: (year: number) => void;
  onToggleWeek: (weekKey: string) => void;
  onToggleProject: (groupKey: string) => void;
  onOpenRun: (runId: string) => void;
  onToggleSelect: (runId: string) => void;
  onRename: (run: BautagebuchRun) => void;
  onDelete: (runId: string) => void;
};

function RunStack({
  runs,
  selectionMode,
  selectedRunIds,
  onOpenRun,
  onToggleSelect,
  onRename,
  onDelete
}: {
  runs: BautagebuchRun[];
  selectionMode: boolean;
  selectedRunIds: string[];
  onOpenRun: (runId: string) => void;
  onToggleSelect: (runId: string) => void;
  onRename: (run: BautagebuchRun) => void;
  onDelete: (runId: string) => void;
}) {
  return (
    <View style={styles.runStack}>
      {runs.map((run) => (
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
      ))}
    </View>
  );
}

function ProjectBlock({
  weekKey,
  project,
  expanded,
  selectionMode,
  selectedRunIds,
  onToggleProject,
  onOpenRun,
  onToggleSelect,
  onRename,
  onDelete
}: {
  weekKey: string;
  project: ProjectRunGroup;
  expanded: boolean;
  selectionMode: boolean;
  selectedRunIds: string[];
  onToggleProject: (groupKey: string) => void;
  onOpenRun: (runId: string) => void;
  onToggleSelect: (runId: string) => void;
  onRename: (run: BautagebuchRun) => void;
  onDelete: (runId: string) => void;
}) {
  const groupKey = projectGroupKey(weekKey, project.projectKey);

  return (
    <View style={styles.projectGroup}>
      <Pressable style={styles.projectHeader} onPress={() => onToggleProject(groupKey)}>
        <Text style={styles.projectTitle}>{project.projectLabel}</Text>
        <Text style={styles.projectMeta}>
          {expanded ? '▾' : '▸'} {formatRunCount(project.runs.length)}
        </Text>
      </Pressable>
      {expanded ? (
        <RunStack
          runs={project.runs}
          selectionMode={selectionMode}
          selectedRunIds={selectedRunIds}
          onOpenRun={onOpenRun}
          onToggleSelect={onToggleSelect}
          onRename={onRename}
          onDelete={onDelete}
        />
      ) : null}
    </View>
  );
}

function WeekBlock({
  week,
  expanded,
  expandedProjects,
  selectionMode,
  selectedRunIds,
  onToggleWeek,
  onToggleProject,
  onOpenRun,
  onToggleSelect,
  onRename,
  onDelete
}: {
  week: WeekRunGroup;
  expanded: boolean;
  expandedProjects: Set<string>;
  selectionMode: boolean;
  selectedRunIds: string[];
  onToggleWeek: (weekKey: string) => void;
  onToggleProject: (groupKey: string) => void;
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
          {expanded ? '▾' : '▸'} {formatRunCount(week.runCount)}
        </Text>
      </Pressable>
      {expanded
        ? week.projects.map((project) => (
            <ProjectBlock
              key={projectGroupKey(week.weekKey, project.projectKey)}
              weekKey={week.weekKey}
              project={project}
              expanded={expandedProjects.has(projectGroupKey(week.weekKey, project.projectKey))}
              selectionMode={selectionMode}
              selectedRunIds={selectedRunIds}
              onToggleProject={onToggleProject}
              onOpenRun={onOpenRun}
              onToggleSelect={onToggleSelect}
              onRename={onRename}
              onDelete={onDelete}
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
  expandedProjects,
  selectionMode,
  selectedRunIds,
  onToggleYear,
  onToggleWeek,
  onToggleProject,
  onOpenRun,
  onToggleSelect,
  onRename,
  onDelete
}: Props) {
  const renderWeeks = (weeks: WeekRunGroup[]) =>
    weeks.map((week) => (
      <WeekBlock
        key={week.weekKey}
        week={week}
        expanded={expandedWeeks.has(week.weekKey)}
        expandedProjects={expandedProjects}
        selectionMode={selectionMode}
        selectedRunIds={selectedRunIds}
        onToggleWeek={onToggleWeek}
        onToggleProject={onToggleProject}
        onOpenRun={onOpenRun}
        onToggleSelect={onToggleSelect}
        onRename={onRename}
        onDelete={onDelete}
      />
    ));

  if (tree.multiYear) {
    return (
      <View style={styles.root}>
        {tree.years.map((yearGroup) => {
          const yearExpanded = expandedYears.has(yearGroup.year);
          return (
            <View key={yearGroup.year} style={styles.yearGroup}>
              <Pressable style={styles.yearHeader} onPress={() => onToggleYear(yearGroup.year)}>
                <Text style={styles.yearTitle}>
                  {yearGroup.year > 0 ? String(yearGroup.year) : 'Ohne Jahr'}
                </Text>
                <Text style={styles.yearMeta}>
                  {yearExpanded ? '▾' : '▸'} {formatRunCount(yearGroup.runCount)}
                </Text>
              </Pressable>
              {yearExpanded ? renderWeeks(yearGroup.weeks) : null}
            </View>
          );
        })}
      </View>
    );
  }

  const weeks = tree.years[0]?.weeks || tree.years.flatMap((year) => year.weeks);

  return <View style={styles.root}>{renderWeeks(weeks)}</View>;
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
  },
  projectGroup: {
    gap: spacing.sm,
    marginLeft: spacing.sm
  },
  projectHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.xxs,
    paddingVertical: spacing.xxs
  },
  projectTitle: {
    ...typography.label,
    color: colors.ink,
    flex: 1
  },
  projectMeta: {
    ...typography.caption,
    color: colors.muted
  },
  runStack: {
    flexDirection: 'column',
    gap: spacing.sm,
    width: '100%'
  }
});
