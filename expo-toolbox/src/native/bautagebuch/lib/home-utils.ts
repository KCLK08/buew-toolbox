import { parseBtbTitle } from './btb-naming';
import type { BautagebuchRun } from '../types';

export type BtbHomeStats = {
  total: number;
  drafts: number;
  completed: number;
};

export function computeBtbHomeStats(runs: BautagebuchRun[]): BtbHomeStats {
  const completed = runs.filter((run) => run.status === 'completed').length;
  return {
    total: runs.length,
    drafts: runs.length - completed,
    completed
  };
}

export function getRecentRuns(runs: BautagebuchRun[], limit = 5): BautagebuchRun[] {
  return [...runs]
    .sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime())
    .slice(0, limit);
}

export function displayRunTitle(title: string): string {
  const parsed = parseBtbTitle(title);
  if (parsed?.input) return parsed.input;
  return String(title || '').trim() || 'Bautagebuch';
}

export function formatActivityTime(iso: string): string {
  const date = new Date(iso);
  const now = new Date();
  const time = date.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });

  if (date.toDateString() === now.toDateString()) {
    return `Heute ${time}`;
  }

  const yesterday = new Date(now);
  yesterday.setDate(yesterday.getDate() - 1);
  if (date.toDateString() === yesterday.toDateString()) {
    return `Gestern ${time}`;
  }

  return date.toLocaleString('de-DE', {
    day: '2-digit',
    month: '2-digit',
    year: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
}
