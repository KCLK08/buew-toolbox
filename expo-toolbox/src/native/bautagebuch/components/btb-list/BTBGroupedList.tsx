import { Pressable, StyleSheet, Text, View } from 'react-native';

import { Card } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { formatGroupLabel } from '../../lib/btb-filter';
import type { ProjectFirstTree, ProjectFirstWeek } from '../../lib/btb-filter';
import {
  projectGroupKey,
  type BautagebuchRunTree,
  type ProjectRunGroup,
  type WeekRunGroup
} from '../../lib/group-runs-by-calendar';
import type { BautagebuchRun } from '../../types';
import { BautagebuchRunCard } from '../BautagebuchRunCard';

type CalendarProps = {
  mode: 'calendar';
  tree: BautagebuchRunTree;
  expandedYears: Set<number>;
  expandedWeeks: Set<string>;
  expandedProjects: Set<string>;
  onToggleYear: (year: number) => void;
  onToggleWeek: (weekKey: string) => void;
  onToggleProject: (groupKey: string) => void;
  onOpenRun: (runId: string) => void;
  onRename: (run: BautagebuchRun) => void;
  onDelete: (runId: string) => void;
};

type ProjectProps = {
  mode: 'project';
  tree: ProjectFirstTree;
  expandedYears: Set<number>;
  expandedWeeks: Set<string>;
  onToggleYear: (year: number) => void;
  onToggleWeek: (weekKey: string) => void;
  onOpenRun: (runId: string) => void;
  onRename: (run: BautagebuchRun) => void;
  onDelete: (runId: string) => void;
};

type Props = CalendarProps | ProjectProps;

function RunStack({
  runs,
  onOpenRun,
  onRename,
  onDelete
}: {
  runs: BautagebuchRun[];
  onOpenRun: (runId: string) => void;
  onRename: (run: BautagebuchRun) => void;
  onDelete: (runId: string) => void;
}) {
  return (
    <View style={styles.runStack}>
      {runs.map((run) => (
        <BautagebuchRunCard
          key={run.runId}
          run={run}
          onPress={() => onOpenRun(run.runId)}
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
  onToggleProject,
  onOpenRun,
  onRename,
  onDelete
}: {
  weekKey: string;
  project: ProjectRunGroup;
  expanded: boolean;
  onToggleProject: (groupKey: string) => void;
  onOpenRun: (runId: string) => void;
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
        <Text style={styles.projectTitle}>{formatGroupLabel(project.projectLabel, project.runs.length)}</Text>
        <Text style={styles.projectMeta}>{expanded ? '▾' : '▸'}</Text>
      </Pressable>
      {expanded ? (
        <View style={styles.projectBody}>
          <RunStack runs={project.runs} onOpenRun={onOpenRun} onRename={onRename} onDelete={onDelete} />
        </View>
      ) : null}
    </View>
  );
}

function WeekBlock({
  week,
  expanded,
  expandedProjects,
  directRuns,
  onToggleWeek,
  onToggleProject,
  onOpenRun,
  onRename,
  onDelete
}: {
  week: Pick<WeekRunGroup, 'weekKey' | 'weekLabel' | 'dateRangeLabel' | 'runCount'> & {
    projects?: ProjectRunGroup[];
  };
  expanded: boolean;
  expandedProjects?: Set<string>;
  directRuns?: BautagebuchRun[];
  onToggleWeek: (weekKey: string) => void;
  onToggleProject?: (groupKey: string) => void;
  onOpenRun: (runId: string) => void;
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
          <Text style={styles.weekTitle}>{formatGroupLabel(week.weekLabel, week.runCount)}</Text>
          <Text style={styles.weekRange}>{week.dateRangeLabel}</Text>
        </View>
        <Text style={styles.weekMeta}>{expanded ? '▾' : '▸'}</Text>
      </Pressable>
      {expanded ? (
        <View style={styles.weekBody}>
          {directRuns ? (
            <RunStack runs={directRuns} onOpenRun={onOpenRun} onRename={onRename} onDelete={onDelete} />
          ) : (
            (week.projects || []).map((project) => (
              <ProjectBlock
                key={projectGroupKey(week.weekKey, project.projectKey)}
                weekKey={week.weekKey}
                project={project}
                expanded={expandedProjects?.has(projectGroupKey(week.weekKey, project.projectKey)) || false}
                onToggleProject={onToggleProject!}
                onOpenRun={onOpenRun}
                onRename={onRename}
                onDelete={onDelete}
              />
            ))
          )}
        </View>
      ) : null}
    </Card>
  );
}

export function BTBGroupedList(props: Props) {
  if (props.mode === 'project') {
    const { tree, expandedYears, expandedWeeks, onToggleYear, onToggleWeek, onOpenRun, onRename, onDelete } = props;

    return (
      <View style={styles.root}>
        <Card style={styles.projectBanner} padded={false}>
          <View style={styles.projectBannerInner}>
            <Text style={styles.projectBannerLabel}>Projekt</Text>
            <Text style={styles.projectBannerTitle}>
              {formatGroupLabel(tree.projectLabel, tree.runCount)}
            </Text>
          </View>
        </Card>

        {tree.years.map((yearGroup) => {
          const yearExpanded = expandedYears.has(yearGroup.year);
          return (
            <Card key={yearGroup.year} style={styles.yearGroup} padded={false}>
              <Pressable
                style={[styles.yearHeader, yearExpanded ? styles.yearHeaderExpanded : null]}
                onPress={() => onToggleYear(yearGroup.year)}
              >
                <Text style={styles.yearTitle}>
                  {formatGroupLabel(yearGroup.year > 0 ? String(yearGroup.year) : 'Ohne Jahr', yearGroup.runCount)}
                </Text>
                <Text style={styles.yearMeta}>{yearExpanded ? '▾' : '▸'}</Text>
              </Pressable>
              {yearExpanded ? (
                <View style={styles.yearBody}>
                  {yearGroup.weeks.map((week: ProjectFirstWeek) => (
                    <WeekBlock
                      key={week.weekKey}
                      week={week}
                      expanded={expandedWeeks.has(week.weekKey)}
                      directRuns={week.runs}
                      onToggleWeek={onToggleWeek}
                      onOpenRun={onOpenRun}
                      onRename={onRename}
                      onDelete={onDelete}
                    />
                  ))}
                </View>
              ) : null}
            </Card>
          );
        })}
      </View>
    );
  }

  const {
    tree,
    expandedYears,
    expandedWeeks,
    expandedProjects,
    onToggleYear,
    onToggleWeek,
    onToggleProject,
    onOpenRun,
    onRename,
    onDelete
  } = props;

  const renderWeeks = (weeks: WeekRunGroup[]) =>
    weeks.map((week) => (
      <WeekBlock
        key={week.weekKey}
        week={week}
        expanded={expandedWeeks.has(week.weekKey)}
        expandedProjects={expandedProjects}
        onToggleWeek={onToggleWeek}
        onToggleProject={onToggleProject}
        onOpenRun={onOpenRun}
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
                  {formatGroupLabel(yearGroup.year > 0 ? String(yearGroup.year) : 'Ohne Jahr', yearGroup.runCount)}
                </Text>
                <Text style={styles.yearMeta}>{yearExpanded ? '▾' : '▸'}</Text>
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
  projectBanner: { overflow: 'hidden' },
  projectBannerInner: {
    paddingHorizontal: spacing.cardPadding,
    paddingVertical: spacing.sm,
    gap: 2
  },
  projectBannerLabel: {
    ...typography.caption,
    color: colors.muted
  },
  projectBannerTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
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
    color: colors.ink,
    flex: 1
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
    color: colors.muted
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
