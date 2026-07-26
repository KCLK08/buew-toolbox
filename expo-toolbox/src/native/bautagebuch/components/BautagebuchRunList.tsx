import { Pressable, StyleSheet, Text, View } from 'react-native';

import { Card } from '../../../components/mobile';
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
      <Pressable
        style={[styles.projectHeader, expanded ? styles.projectHeaderExpanded : null]}
        onPress={() => onToggleProject(groupKey)}
      >
        <Text style={styles.projectTitle}>{project.projectLabel}</Text>
        <Text style={styles.projectMeta}>
          {expanded ? '▾' : '▸'} {formatRunCount(project.runs.length)}
        </Text>
      </Pressable>
      {expanded ? (
        <View style={styles.projectBody}>
          <RunStack
            runs={project.runs}
            selectionMode={selectionMode}
            selectedRunIds={selectedRunIds}
            onOpenRun={onOpenRun}
            onToggleSelect={onToggleSelect}
            onRename={onRename}
            onDelete={onDelete}
          />
        </View>
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
    <Card style={styles.weekGroup} padded={false}>
      <Pressable
        style={[styles.weekHeader, expanded ? styles.weekHeaderExpanded : null]}
        onPress={() => onToggleWeek(week.weekKey)}
      >
        <View style={styles.weekHeaderMain}>
          <Text style={styles.weekTitle}>{week.weekLabel}</Text>
          <Text style={styles.weekRange}>{week.dateRangeLabel}</Text>
        </View>
        <Text style={styles.weekMeta}>
          {expanded ? '▾' : '▸'} {formatRunCount(week.runCount)}
        </Text>
      </Pressable>
      {expanded ? (
        <View style={styles.weekBody}>
          {week.projects.map((project) => (
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
          ))}
        </View>
      ) : null}
    </Card>
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
            <Card key={yearGroup.year} style={styles.yearGroup} padded={false}>
              <Pressable
                style={[styles.yearHeader, yearExpanded ? styles.yearHeaderExpanded : null]}
                onPress={() => onToggleYear(yearGroup.year)}
              >
                <Text style={styles.yearTitle}>
                  {yearGroup.year > 0 ? String(yearGroup.year) : 'Ohne Jahr'}
                </Text>
                <Text style={styles.yearMeta}>
                  {yearExpanded ? '▾' : '▸'} {formatRunCount(yearGroup.runCount)}
                </Text>
              </Pressable>
              {yearExpanded ? <View style={styles.yearBody}>{renderWeeks(yearGroup.weeks)}</View> : null}
            </Card>
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
  yearGroup: { overflow: 'hidden' },
  yearHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.cardPadding,
    paddingVertical: spacing.sm
  },
  yearHeaderExpanded: {
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  yearBody: {
    gap: spacing.sm,
    padding: spacing.sm
  },
  yearTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  yearMeta: {
    ...typography.caption,
    color: colors.muted
  },
  weekGroup: { overflow: 'hidden' },
  weekHeader: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    gap: spacing.sm,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.cardPadding,
    paddingVertical: spacing.sm
  },
  weekHeaderExpanded: {
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  weekBody: {
    gap: spacing.sm,
    padding: spacing.sm,
    backgroundColor: colors.bg
  },
  weekHeaderMain: {
    flex: 1,
    gap: 2
  },
  weekTitle: {
    ...typography.subtitle,
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
    overflow: 'hidden',
    borderRadius: spacing.inputRadius,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated
  },
  projectHeader: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm,
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.xs
  },
  projectHeaderExpanded: {
    borderBottomWidth: 1,
    borderBottomColor: colors.border
  },
  projectBody: {
    padding: spacing.sm,
    gap: spacing.sm
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
    gap: spacing.xs,
    width: '100%'
  }
});
