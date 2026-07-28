import {
  formatRunCount,
  groupRunsByCalendar,
  resolveProjectFieldKey,
  resolveProjectLabel,
  type BautagebuchRunTree,
  type ProjectRunGroup,
  type WeekRunGroup,
  type YearRunGroup
} from './group-runs-by-calendar';
import type { BautagebuchRun } from '../types';

export type BtbSortOrder = 'newest' | 'oldest';
export type BtbGroupMode = 'calendar' | 'project';

export type BtbListFilters = {
  groupMode: BtbGroupMode;
  projectKey: string | null;
  sortOrder: BtbSortOrder;
};

export type ProjectListItem = {
  projectKey: string;
  projectLabel: string;
  runCount: number;
};

export type ProjectFirstWeek = {
  weekKey: string;
  weekYear: number;
  weekNumber: number;
  weekLabel: string;
  dateRangeLabel: string;
  runCount: number;
  runs: BautagebuchRun[];
};

export type ProjectFirstYear = {
  year: number;
  runCount: number;
  weeks: ProjectFirstWeek[];
};

export type ProjectFirstTree = {
  projectKey: string;
  projectLabel: string;
  runCount: number;
  years: ProjectFirstYear[];
};

export const DEFAULT_BTB_FILTERS: BtbListFilters = {
  groupMode: 'project',
  projectKey: null,
  sortOrder: 'newest'
};

function projectKeyFromLabel(label: string): string {
  return label
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[^\wäöüß\- ]+/gi, '')
    .slice(0, 80);
}

function sortRuns(runs: BautagebuchRun[], sortOrder: BtbSortOrder): BautagebuchRun[] {
  return [...runs].sort((left, right) => {
    const leftTs = new Date(left.updatedAt || left.createdAt).getTime();
    const rightTs = new Date(right.updatedAt || right.createdAt).getTime();
    return sortOrder === 'newest' ? rightTs - leftTs : leftTs - rightTs;
  });
}

export function filterRunsByProject(
  runs: BautagebuchRun[],
  setupModel: Record<string, unknown> | null | undefined,
  projectKey: string | null
): BautagebuchRun[] {
  if (!projectKey) return runs;
  const projectFieldKey = resolveProjectFieldKey(setupModel);
  return runs.filter(
    (run) =>
      (projectKeyFromLabel(resolveProjectLabel(run, projectFieldKey)) || 'ohne-projekt') === projectKey
  );
}

export function resolveProjectLabelByKey(
  projectKey: string | null,
  projects: ProjectListItem[]
): string | null {
  if (!projectKey) return null;
  return projects.find((entry) => entry.projectKey === projectKey)?.projectLabel ?? null;
}

export function listProjectsFromRuns(
  runs: BautagebuchRun[],
  setupModel?: Record<string, unknown> | null
): ProjectListItem[] {
  const projectFieldKey = resolveProjectFieldKey(setupModel);
  const map = new Map<string, ProjectListItem>();

  for (const run of runs) {
    const projectLabel = resolveProjectLabel(run, projectFieldKey);
    const projectKey = projectKeyFromLabel(projectLabel) || 'ohne-projekt';
    const existing = map.get(projectKey);
    if (existing) {
      existing.runCount += 1;
      continue;
    }
    map.set(projectKey, { projectKey, projectLabel, runCount: 1 });
  }

  return [...map.values()].sort((a, b) => a.projectLabel.localeCompare(b.projectLabel, 'de'));
}

function flattenWeekRuns(week: WeekRunGroup): BautagebuchRun[] {
  return week.projects.flatMap((project) => project.runs);
}

export function buildCalendarTree(
  runs: BautagebuchRun[],
  setupModel: Record<string, unknown> | null | undefined,
  filters: Pick<BtbListFilters, 'sortOrder'>
): BautagebuchRunTree {
  const sorted = sortRuns(runs, filters.sortOrder);
  return groupRunsByCalendar(sorted, { setupModel });
}

export function buildProjectFirstTree(
  runs: BautagebuchRun[],
  setupModel: Record<string, unknown> | null | undefined,
  projectKey: string,
  filters: Pick<BtbListFilters, 'sortOrder'>
): ProjectFirstTree | null {
  const projects = listProjectsFromRuns(runs, setupModel);
  const project = projects.find((entry) => entry.projectKey === projectKey);
  if (!project) return null;

  const projectFieldKey = resolveProjectFieldKey(setupModel);
  const projectRuns = sortRuns(
    runs.filter((run) => projectKeyFromLabel(resolveProjectLabel(run, projectFieldKey)) === projectKey),
    filters.sortOrder
  );

  const tree = groupRunsByCalendar(projectRuns, { setupModel });
  const years: ProjectFirstYear[] = tree.years.map((year) => ({
    year: year.year,
    runCount: year.runCount,
    weeks: year.weeks.map((week) => ({
      weekKey: week.weekKey,
      weekYear: week.weekYear,
      weekNumber: week.weekNumber,
      weekLabel: week.weekLabel,
      dateRangeLabel: week.dateRangeLabel,
      runCount: week.runCount,
      runs: sortRuns(flattenWeekRuns(week), filters.sortOrder)
    }))
  }));

  const runCount = years.reduce((sum, year) => sum + year.runCount, 0);

  return {
    projectKey: project.projectKey,
    projectLabel: project.projectLabel,
    runCount,
    years
  };
}

export function formatGroupLabel(label: string, count: number): string {
  return `${label} (${formatRunCount(count)})`;
}

export function countFilteredRuns(tree: BautagebuchRunTree): number {
  return tree.years.reduce(
    (sum, year) => sum + year.weeks.reduce((weekSum, week) => weekSum + week.runCount, 0),
    0
  );
}

export type { BautagebuchRunTree, ProjectRunGroup, WeekRunGroup, YearRunGroup };
