import type { BautagebuchRun } from '../types';

function valueToSearchText(value: unknown): string {
  if (value == null) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (Array.isArray(value)) return value.map(valueToSearchText).join(' ');
  if (typeof value === 'object') {
    return Object.values(value as Record<string, unknown>).map(valueToSearchText).join(' ');
  }
  return '';
}

export function runSearchText(run: BautagebuchRun): string {
  const parts = [
    run.title,
    ...Object.values(run.values || {}).map(valueToSearchText),
    ...(run.photoDoc?.entries || []).flatMap((entry) => [entry.id, entry.localPath || ''])
  ];
  return parts.join(' ').toLowerCase();
}

export function runMatchesQuery(run: BautagebuchRun, query: string): boolean {
  const normalized = query.trim().toLowerCase();
  if (!normalized) return true;
  return runSearchText(run).includes(normalized);
}

export function filterRunsBySearchQuery(runs: BautagebuchRun[], query: string): BautagebuchRun[] {
  const normalized = query.trim();
  if (!normalized) return runs;
  return runs.filter((run) => runMatchesQuery(run, normalized));
}
