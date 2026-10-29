import { parseBtbTitle } from './btb-naming.js';
import { parseBtbDateFromRun } from './btb-date.js';
import { inputKeyForField } from './setup-model.js';

const DAY_MS = 86_400_000;

/** @typedef {{ runId: string; title?: string; values?: Record<string, unknown>; createdAt?: string; updatedAt?: string }} BautagebuchRun */

/** @typedef {{ projectKey: string; projectLabel: string; runs: BautagebuchRun[] }} ProjectRunGroup */

/** @typedef {{ weekKey: string; weekYear: number; weekNumber: number; weekLabel: string; dateRangeLabel: string; projects: ProjectRunGroup[]; runCount: number }} WeekRunGroup */

/** @typedef {{ year: number; weeks: WeekRunGroup[]; runCount: number }} YearRunGroup */

/** @typedef {{ multiYear: boolean; years: YearRunGroup[] }} BautagebuchRunTree */

function compareBySortOrder(leftValue, rightValue, sortOrder = 'newest') {
  return sortOrder === 'newest' ? rightValue - leftValue : leftValue - rightValue;
}

function parseBtbDate(run) {
  return parseBtbDateFromRun(run);
}

export function resolveProjectFieldKey(setupModel) {
  const sections =
    (setupModel?.single_sections) || [];
  for (const section of sections) {
    for (const field of section.fields || []) {
      if (String(field.fieldName || '').trim() === 'Text1') {
        const fieldId = String(field.fieldId || '').trim();
        if (!fieldId) continue;
        return inputKeyForField({ fieldId });
      }
    }
  }
  return null;
}

export function resolveProjectLabel(run, projectFieldKey) {
  if (projectFieldKey) {
    const fromValues = String(run.values?.[projectFieldKey] ?? '').trim();
    if (fromValues) return fromValues;
  }

  const parsed = parseBtbTitle(String(run.title || ''));
  if (parsed?.input) return parsed.input;

  const title = String(run.title || '').trim();
  if (title) return title;
  return 'Ohne Projekt';
}

function projectKeyFromLabel(label) {
  return label
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[^\wäöüß\- ]+/gi, '')
    .slice(0, 80);
}

function isoWeekMeta(date) {
  const target = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  const weekday = target.getUTCDay() || 7;
  target.setUTCDate(target.getUTCDate() + 4 - weekday);
  const weekYear = target.getUTCFullYear();
  const yearStart = new Date(Date.UTC(weekYear, 0, 1));
  const weekNumber = Math.ceil(((target.getTime() - yearStart.getTime()) / DAY_MS + 1) / 7);
  return { weekYear, weekNumber };
}

function isoWeekDateRange(weekYear, weekNumber) {
  const jan4 = new Date(Date.UTC(weekYear, 0, 4));
  const jan4Day = jan4.getUTCDay() || 7;
  const mondayWeek1 = new Date(jan4);
  mondayWeek1.setUTCDate(jan4.getUTCDate() - jan4Day + 1);
  const monday = new Date(mondayWeek1);
  monday.setUTCDate(mondayWeek1.getUTCDate() + (weekNumber - 1) * 7);
  const sunday = new Date(monday);
  sunday.setUTCDate(monday.getUTCDate() + 6);
  return { start: monday, end: sunday };
}

function formatWeekDateRange(start, end) {
  const dayMonth = { day: '2-digit', month: '2-digit' };
  const startLabel = start.toLocaleDateString('de-DE', dayMonth);
  const endLabel = end.toLocaleDateString('de-DE', { ...dayMonth, year: 'numeric' });
  return `${startLabel} – ${endLabel}`;
}

function sortRunsByDate(runs, sortOrder = 'newest') {
  return [...runs].sort((left, right) => {
    const leftDate = parseBtbDate(left)?.getTime() ?? Number.NEGATIVE_INFINITY;
    const rightDate = parseBtbDate(right)?.getTime() ?? Number.NEGATIVE_INFINITY;
    if (leftDate !== rightDate) return compareBySortOrder(leftDate, rightDate, sortOrder);
    const leftUpdated = String(left.updatedAt || left.createdAt || '');
    const rightUpdated = String(right.updatedAt || right.createdAt || '');
    return sortOrder === 'newest'
      ? rightUpdated.localeCompare(leftUpdated)
      : leftUpdated.localeCompare(rightUpdated);
  });
}

function groupRunsByProject(runs, projectFieldKey, sortOrder = 'newest') {
  const projectMap = new Map();

  for (const run of runs) {
    const projectLabel = resolveProjectLabel(run, projectFieldKey);
    const projectKey = projectKeyFromLabel(projectLabel) || 'ohne-projekt';
    const bucket = projectMap.get(projectKey);
    if (bucket) {
      bucket.runs.push(run);
      continue;
    }
    projectMap.set(projectKey, {
      projectKey,
      projectLabel,
      runs: [run]
    });
  }

  return [...projectMap.values()]
    .map((group) => ({
      ...group,
      runs: sortRunsByDate(group.runs, sortOrder)
    }))
    .sort((left, right) => {
      const leftDate = parseBtbDate(left.runs[0])?.getTime() ?? Number.NEGATIVE_INFINITY;
      const rightDate = parseBtbDate(right.runs[0])?.getTime() ?? Number.NEGATIVE_INFINITY;
      if (leftDate !== rightDate) return compareBySortOrder(leftDate, rightDate, sortOrder);
      return left.projectLabel.localeCompare(right.projectLabel, 'de');
    });
}

export function groupRunsByCalendar(runs, options = {}) {
  const sortOrder = options.sortOrder || 'newest';
  const projectFieldKey = resolveProjectFieldKey(options.setupModel);
  const weekMap = new Map();

  for (const run of runs) {
    const btbDate = parseBtbDate(run);
    let weekKey = 'unknown';
    if (btbDate) {
      const { weekYear, weekNumber } = isoWeekMeta(btbDate);
      weekKey = `${weekYear}-${String(weekNumber).padStart(2, '0')}`;
    }
    const bucket = weekMap.get(weekKey) || [];
    bucket.push(run);
    weekMap.set(weekKey, bucket);
  }

  const weeks = [...weekMap.entries()].map(([weekKey, weekRuns]) => {
    const projects = groupRunsByProject(weekRuns, projectFieldKey, sortOrder);
    const runCount = weekRuns.length;

    if (weekKey === 'unknown') {
      return {
        weekKey,
        weekYear: 0,
        weekNumber: 0,
        weekLabel: 'Ohne Datum',
        dateRangeLabel: 'Datum nicht erkannt',
        projects,
        runCount
      };
    }

    const [weekYearRaw, weekNumberRaw] = weekKey.split('-');
    const weekYear = Number(weekYearRaw);
    const weekNumber = Number(weekNumberRaw);
    const { start, end } = isoWeekDateRange(weekYear, weekNumber);

    return {
      weekKey,
      weekYear,
      weekNumber,
      weekLabel: `KW ${String(weekNumber).padStart(2, '0')}`,
      dateRangeLabel: formatWeekDateRange(start, end),
      projects,
      runCount
    };
  });

  weeks.sort((left, right) => {
    if (left.weekYear !== right.weekYear) {
      return compareBySortOrder(left.weekYear, right.weekYear, sortOrder);
    }
    return compareBySortOrder(left.weekNumber, right.weekNumber, sortOrder);
  });

  const yearMap = new Map();
  for (const week of weeks) {
    const year = week.weekYear || 0;
    const bucket = yearMap.get(year) || [];
    bucket.push(week);
    yearMap.set(year, bucket);
  }

  const years = [...yearMap.entries()]
    .map(([year, yearWeeks]) => ({
      year,
      weeks: yearWeeks,
      runCount: yearWeeks.reduce((sum, week) => sum + week.runCount, 0)
    }))
    .sort((left, right) => compareBySortOrder(left.year, right.year, sortOrder));

  const distinctYears = years.filter((entry) => entry.year > 0);

  return {
    multiYear: distinctYears.length > 1,
    years
  };
}

export function formatRunCount(count) {
  return `${count} BTB${count === 1 ? '' : 's'}`;
}

export function projectGroupKey(weekKey, projectKey) {
  return `${weekKey}:${projectKey}`;
}
