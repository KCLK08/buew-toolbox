import type { DefectPriority, EntityStatus } from '../types/offline';

export function formatRelativeDate(iso: string | null | undefined): string {
  if (!iso) return '—';
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return iso;
  const now = new Date();
  const startToday = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const startThat = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
  const dayDiff = Math.round((startToday - startThat) / (24 * 60 * 60 * 1000));
  if (dayDiff === 0) return 'Heute';
  if (dayDiff === 1) return 'Gestern';
  return date.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

export function formatStatusLabel(status: EntityStatus | string): string {
  switch (status) {
    case 'active':
      return 'Aktiv';
    case 'draft':
      return 'Entwurf';
    case 'archived':
      return 'Archiv';
    case 'completed':
      return 'Abgeschlossen';
    default:
      return String(status);
  }
}

export function statusTone(status: EntityStatus | string): 'success' | 'warning' | 'neutral' | 'info' | 'accent' {
  switch (status) {
    case 'active':
      return 'success';
    case 'completed':
      return 'info';
    case 'draft':
      return 'warning';
    case 'archived':
      return 'neutral';
    default:
      return 'accent';
  }
}

export function priorityLabel(priority: DefectPriority | string): string {
  switch (priority) {
    case 'low':
      return 'Niedrig';
    case 'normal':
      return 'Normal';
    case 'high':
      return 'Hoch';
    case 'critical':
      return 'Kritisch';
    default:
      return String(priority);
  }
}

export function priorityTone(priority: DefectPriority | string): 'neutral' | 'warning' | 'danger' | 'info' {
  switch (priority) {
    case 'low':
      return 'info';
    case 'normal':
      return 'neutral';
    case 'high':
      return 'warning';
    case 'critical':
      return 'danger';
    default:
      return 'neutral';
  }
}
