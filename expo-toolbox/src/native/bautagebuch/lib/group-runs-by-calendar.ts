import { inputKeyForField } from './setup-model.js';
import type { BautagebuchRun } from '../types';

const DAY_MS = 86_400_000;

export type ProjectRunGroup = {
  projectKey: string;
  projectLabel: string;
  runs: BautagebuchRun[];
};

export type WeekRunGroup = {
  weekKey: string;
  weekYear: number;
  weekNumber: number;
  weekLabel: string;
  dateRangeLabel: string;
  projects: ProjectRunGroup[];
  runCount: number;
};

export type YearRunGroup = {
  year: number;
  weeks: WeekRunGroup[];
  runCount: number;
};

export type BautagebuchRunTree = {
  multiYear: boolean;
  years: YearRunGroup[];
};

type GroupingOptions = {
  setupModel?: Record<string, unknown> | null;
};

function parseBtbDate(run: BautagebuchRun): Date | null {
  const fromTitle = run.title.match(/\d{4}-\d{2}-\d{2}/)?.[0];
  if (fromTitle) {
    return new Date(`${fromTitle}T12:00:00`);
  }
  const fallback = run.createdAt?.slice(0, 10);
  if (fallback) {
    return new Date(`${fallback}T12:00:00`);
  }
  return null;
}

function resolveProjectFieldKey(setupModel?: Record<string, unknown> | null): string | null {
  const sections =
    (setupModel?.single_sections as Array<{ fields?: Array<{ fieldId?: string; fieldName?: string }> }>) || [];
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

export function resolveProjectLabel(
  run: BautagebuchRun,
  projectFieldKey: string | null
): string {
  if (projectFieldKey) {
    const fromValues = String(run.values?.[projectFieldKey] ?? '').trim();
    if (fromValues) return fromValues;
  }

  const title = String(run.title || '').trim();
  const match = title.match(/^BTB\s+\d{4}-\d{2}-\d{2}\s*-\s*(.+)$/i);
  if (match?.[1]?.trim()) return match[1].trim();
  if (title) return title;
  return 'Ohne Projekt';
}

function projectKeyFromLabel(label: string): string {
  return label
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[^\wäöüß\- ]+/gi, '')
    .slice(0, 80);
}

function isoWeekMeta(date: Date): { weekYear: number; weekNumber: number } {
  const target = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const weekday = target.getUTCDay() || 7;
  target.setUTCDate(target.getUTCDate() + 4 - weekday);
  const weekYear = target.getUTCFullYear();
  const yearStart = new Date(Date.UTC(weekYear, 0, 1));
  const weekNumber = Math.ceil(((target.getTime() - yearStart.getTime()) / DAY_MS + 1) / 7);
  return { weekYear, weekNumber };
}

function isoWeekDateRange(weekYear: number, weekNumber: number): { start: Date; end: Date } {
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

function formatWeekDateRange(start: Date, end: Date): string {
  const dayMonth: Intl.DateTimeFormatOptions = { day: '2-digit', month: '2-digit' };
  const startLabel = start.toLocaleDateString('de-DE', dayMonth);
  const endLabel = end.toLocaleDateString('de-DE', { ...dayMonth, year: 'numeric' });
  return `${startLabel} – ${endLabel}`;
}

function sortRunsByDateDesc(runs: BautagebuchRun[]): BautagebuchRun[] {
  return [...runs].sort((left, right) => {
    const leftDate = parseBtbDate(left)?.getTime() ?? Number.NEGATIVE_INFINITY;
    const rightDate = parseBtbDate(right)?.getTime() ?? Number.NEGATIVE_INFINITY;
    if (leftDate !== rightDate) return rightDate - leftDate;
    return String(right.updatedAt || '').localeCompare(String(left.updatedAt || ''));
  });
}

function groupRunsByProject(
  runs: BautagebuchRun[],
  projectFieldKey: string | null
): ProjectRunGroup[] {
  const projectMap = new Map<string, ProjectRunGroup>();

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
      runs: sortRunsByDateDesc(group.runs)
    }))
    .sort((left, right) => {
      const leftDate = parseBtbDate(left.runs[0])?.getTime() ?? Number.NEGATIVE_INFINITY;
      const rightDate = parseBtbDate(right.runs[0])?.getTime() ?? Number.NEGATIVE_INFINITY;
      if (leftDate !== rightDate) return rightDate - leftDate;
      return left.projectLabel.localeCompare(right.projectLabel, 'de');
    });
}

export function groupRunsByCalendar(
  runs: BautagebuchRun[],
  options: GroupingOptions = {}
): BautagebuchRunTree {
  const projectFieldKey = resolveProjectFieldKey(options.setupModel);
  const weekMap = new Map<string, BautagebuchRun[]>();

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

  const weeks: WeekRunGroup[] = [...weekMap.entries()].map(([weekKey, weekRuns]) => {
    const projects = groupRunsByProject(weekRuns, projectFieldKey);
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
    if (left.weekYear !== right.weekYear) return right.weekYear - left.weekYear;
    return right.weekNumber - left.weekNumber;
  });

  const yearMap = new Map<number, WeekRunGroup[]>();
  for (const week of weeks) {
    const year = week.weekYear || 0;
    const bucket = yearMap.get(year) || [];
    bucket.push(week);
    yearMap.set(year, bucket);
  }

  const years: YearRunGroup[] = [...yearMap.entries()]
    .map(([year, yearWeeks]) => ({
      year,
      weeks: yearWeeks,
      runCount: yearWeeks.reduce((sum, week) => sum + week.runCount, 0)
    }))
    .sort((left, right) => right.year - left.year);

  const distinctYears = years.filter((entry) => entry.year > 0);

  return {
    multiYear: distinctYears.length > 1,
    years
  };
}

export function formatRunCount(count: number): string {
  return `${count} BTB${count === 1 ? '' : 's'}`;
}

export function projectGroupKey(weekKey: string, projectKey: string): string {
  return `${weekKey}:${projectKey}`;
}
