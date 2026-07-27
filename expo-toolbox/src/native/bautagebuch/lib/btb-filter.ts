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
  year: number | null;
  weekKey: string | null;
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
  groupMode: 'calendar',
  projectKey: null,
  year: null,
  weekKey: null,
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

function filterCalendarTree(tree: BautagebuchRunTree, filters: BtbListFilters): BautagebuchRunTree {
  let years = tree.years;

  if (filters.projectKey) {
    years = years
      .map((year) => ({
        ...year,
        weeks: year.weeks
          .map((week) => ({
            ...week,
            projects: week.projects.filter((project) => project.projectKey === filters.projectKey),
            runCount: week.projects
              .filter((project) => project.projectKey === filters.projectKey)
              .reduce((sum, project) => sum + project.runs.length, 0)
          }))
          .filter((week) => week.runCount > 0),
        runCount: 0
      }))
      .map((year) => ({
        ...year,
        runCount: year.weeks.reduce((sum, week) => sum + week.runCount, 0)
      }))
      .filter((year) => year.runCount > 0);
  }

  if (filters.year !== null) {
    years = years.filter((year) => year.year === filters.year);
  }

  if (filters.weekKey) {
    years = years
      .map((year) => ({
        ...year,
        weeks: year.weeks.filter((week) => week.weekKey === filters.weekKey),
        runCount: year.weeks
          .filter((week) => week.weekKey === filters.weekKey)
          .reduce((sum, week) => sum + week.runCount, 0)
      }))
      .filter((year) => year.weeks.length > 0);
  }

  const distinctYears = years.filter((entry) => entry.year > 0);
  return {
    multiYear: distinctYears.length > 1,
    years
  };
}

export function buildCalendarTree(
  runs: BautagebuchRun[],
  setupModel: Record<string, unknown> | null | undefined,
  filters: BtbListFilters
): BautagebuchRunTree {
  const sorted = sortRuns(runs, filters.sortOrder);
  const tree = groupRunsByCalendar(sorted, { setupModel });
  return filterCalendarTree(tree, filters);
}

export function buildProjectFirstTree(
  runs: BautagebuchRun[],
  setupModel: Record<string, unknown> | null | undefined,
  projectKey: string,
  filters: Pick<BtbListFilters, 'year' | 'weekKey' | 'sortOrder'>
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
  let years: ProjectFirstYear[] = tree.years.map((year) => ({
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

  if (filters.year !== null) {
    years = years.filter((year) => year.year === filters.year);
  }
  if (filters.weekKey) {
    years = years
      .map((year) => ({
        ...year,
        weeks: year.weeks.filter((week) => week.weekKey === filters.weekKey),
        runCount: year.weeks
          .filter((week) => week.weekKey === filters.weekKey)
          .reduce((sum, week) => sum + week.runCount, 0)
      }))
      .filter((year) => year.weeks.length > 0);
  }

  const runCount = years.reduce((sum, year) => sum + year.runCount, 0);

  return {
    projectKey: project.projectKey,
    projectLabel: project.projectLabel,
    runCount,
    years
  };
}

export function listAvailableYears(tree: BautagebuchRunTree): number[] {
  return tree.years.map((year) => year.year).filter((year) => year > 0);
}

export function listAvailableWeeks(tree: BautagebuchRunTree, year: number | null): WeekRunGroup[] {
  const weeks = year === null ? tree.years.flatMap((entry) => entry.weeks) : tree.years.find((entry) => entry.year === year)?.weeks || [];
  return weeks;
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
