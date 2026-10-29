import {
  formatRunCount,
  groupRunsByCalendar,
  resolveProjectFieldKey,
  resolveProjectLabel
} from './group-runs-by-calendar.js';

/** @typedef {{ groupMode: 'calendar' | 'project'; projectKey: string | null; sortOrder: 'newest' | 'oldest' }} BtbListFilters */

/** @typedef {{ projectKey: string; projectLabel: string; runCount: number }} ProjectListItem */

export const DEFAULT_BTB_FILTERS = {
  groupMode: 'project',
  projectKey: null,
  sortOrder: 'newest'
};

function projectKeyFromLabel(label) {
  return label
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[^\wäöüß\- ]+/gi, '')
    .slice(0, 80);
}

function sortRuns(runs, sortOrder) {
  return [...runs].sort((left, right) => {
    const leftTs = new Date(left.updatedAt || left.createdAt).getTime();
    const rightTs = new Date(right.updatedAt || right.createdAt).getTime();
    return sortOrder === 'newest' ? rightTs - leftTs : leftTs - rightTs;
  });
}

export function filterRunsByProject(runs, setupModel, projectKey) {
  if (!projectKey) return runs;
  const projectFieldKey = resolveProjectFieldKey(setupModel);
  return runs.filter(
    (run) =>
      (projectKeyFromLabel(resolveProjectLabel(run, projectFieldKey)) || 'ohne-projekt') === projectKey
  );
}

export function listProjectsFromRuns(runs, setupModel) {
  const projectFieldKey = resolveProjectFieldKey(setupModel);
  const map = new Map();

  for (const run of runs) {
    const projectLabel = resolveProjectLabel(run, projectFieldKey);
    const key = projectKeyFromLabel(projectLabel) || 'ohne-projekt';
    const existing = map.get(key);
    if (existing) {
      existing.runCount += 1;
      continue;
    }
    map.set(key, { projectKey: key, projectLabel, runCount: 1 });
  }

  return [...map.values()].sort((a, b) => a.projectLabel.localeCompare(b.projectLabel, 'de'));
}

function flattenWeekRuns(week) {
  return week.projects.flatMap((project) => project.runs);
}

export function buildCalendarTree(runs, setupModel, filters) {
  const sorted = sortRuns(runs, filters.sortOrder);
  return groupRunsByCalendar(sorted, { setupModel, sortOrder: filters.sortOrder });
}

export function buildProjectFirstTree(runs, setupModel, projectKey, filters) {
  const projects = listProjectsFromRuns(runs, setupModel);
  const project = projects.find((entry) => entry.projectKey === projectKey);
  if (!project) return null;

  const projectFieldKey = resolveProjectFieldKey(setupModel);
  const projectRuns = sortRuns(
    runs.filter((run) => projectKeyFromLabel(resolveProjectLabel(run, projectFieldKey)) === projectKey),
    filters.sortOrder
  );

  const tree = groupRunsByCalendar(projectRuns, { setupModel, sortOrder: filters.sortOrder });
  const years = tree.years.map((year) => ({
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

export function formatGroupLabel(label, count) {
  return `${label} (${formatRunCount(count)})`;
}
