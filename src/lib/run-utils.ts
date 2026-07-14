import type { PhotoDoc, RunSection, SetupModel } from '@/types';
import { buildRunSections, inputKeyForField } from './setup-model';

export const PHOTO_DOC_SECTION_ID = 'photo-doc';
export const PHOTO_DOC_ENABLED_RUN_KEY = '__photoDoc:enabled';
export const TABLE_ROW_COUNT_KEY_PREFIX = '__tableRows:';
export const MAIN_PERSONAL_TABLE_ID = 'table_main_personal';
export const LEISTUNGSBLOCK_TABLE_ID = 'table_detail_blocks';

const RUN_DEFAULTS_BY_FIELD_NAME = new Map([['Text2', 'Kazim Celik']]);
const DATE_FIELD_NAME_PATTERN = /^Date\d+$/i;
const SHIFT_FIELD_NAME_SET = new Set(['Check Box1', 'Check Box2', 'Check Box3']);

export function runSectionOrderRank(section: RunSection) {
  const sectionId = String(section?.sectionId || '').trim();
  const label = String(section?.label || '').trim().toLowerCase();

  if (sectionId === 'single:header') return 10;
  if (sectionId === 'single:weather') return 20;
  if (sectionId === `table:${MAIN_PERSONAL_TABLE_ID}`) return 30;
  if (sectionId === `table:${LEISTUNGSBLOCK_TABLE_ID}`) return 40;
  if (sectionId === 'single:closing') return 50;
  if (sectionId === PHOTO_DOC_SECTION_ID) return 60;

  if (label.includes('kopfdaten')) return 10;
  if (label.includes('witterung')) return 20;
  if (label.includes('firmen') && label.includes('personal')) return 30;
  if (label.includes('leistungsblock') || label.includes('ausgeführte arbeiten')) return 40;
  if (label.includes('abschluss')) return 50;
  return 900;
}

export function orderRunSections(sections: RunSection[] = []) {
  return [...sections]
    .map((section, index) => ({ section, index, rank: runSectionOrderRank(section) }))
    .sort((a, b) => a.rank - b.rank || a.index - b.index)
    .map((entry) => entry.section);
}

export function buildRunSectionsWithPhotoDoc(model: SetupModel, photoDocEnabled: boolean | null): RunSection[] {
  const sections = orderRunSections(buildRunSections(model) as RunSection[]);
  sections.push({
    sectionId: PHOTO_DOC_SECTION_ID,
    kind: 'photo-doc',
    label: 'Fotodokumentation',
  });
  return sections;
}

export function normalizeRunPhotoDoc(value: Partial<PhotoDoc> | null | undefined): PhotoDoc {
  const entries = Array.isArray(value?.entries)
    ? value.entries
        .map((entry) => ({
          id: String(entry?.id || '').trim(),
          createdAt: String(entry?.createdAt || '').trim() || new Date().toISOString(),
          mimeType: String(entry?.mimeType || 'image/jpeg').trim() || 'image/jpeg',
          photoUri: String(entry?.photoUri || '').trim(),
        }))
        .filter((entry) => entry.id && entry.photoUri)
    : [];

  return {
    enabled: value?.enabled ?? null,
    entries,
    updatedAt: String(value?.updatedAt || '').trim() || new Date().toISOString(),
  };
}

export function isPhotoDocEnabled(photoDoc: PhotoDoc, values: Record<string, unknown>): boolean {
  if (photoDoc.enabled === true) return true;
  if (photoDoc.enabled === false) return false;
  return values[PHOTO_DOC_ENABLED_RUN_KEY] === true;
}

export function applyRunDefaultsFromModel(model: SetupModel, values: Record<string, unknown> = {}) {
  const next = { ...values };
  let changed = false;
  const today = new Date().toISOString().slice(0, 10);

  for (const section of model.single_sections || []) {
    for (const field of section.fields || []) {
      if (field.skipped) continue;
      const key = inputKeyForField(field);
      if (next[key] !== undefined && String(next[key] ?? '').trim() !== '') continue;

      const fieldName = String(field.fieldName || '').trim();
      if (DATE_FIELD_NAME_PATTERN.test(fieldName)) {
        next[key] = today;
        changed = true;
        continue;
      }
      if (RUN_DEFAULTS_BY_FIELD_NAME.has(fieldName)) {
        next[key] = RUN_DEFAULTS_BY_FIELD_NAME.get(fieldName) || '';
        changed = true;
      }
    }
  }

  return { values: next, changed };
}

export function buildRunTitleByConvention(name: string): string {
  const normalized = String(name || '').trim();
  if (!normalized) return 'BTB';
  const date = new Date().toISOString().slice(0, 10);
  return `${normalized} – ${date}`;
}

export function normalizeRunNameInput(value: string): string {
  return String(value || '').trim();
}

export function tableRowCountKey(tableId: string) {
  return `${TABLE_ROW_COUNT_KEY_PREFIX}${tableId}`;
}

export function getVisibleRowCount(tableId: string, values: Record<string, unknown>, maxRows: number) {
  const raw = Number(values[tableRowCountKey(tableId)]);
  if (!Number.isFinite(raw)) return 1;
  return Math.max(1, Math.min(maxRows, Math.floor(raw)));
}

export function sanitizeFileName(value: string) {
  const cleaned = String(value || 'bautagebuch')
    .trim()
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_')
    .replace(/\s+/g, '_');
  return cleaned || 'bautagebuch';
}

export function getWeekKey(isoDate: string) {
  const date = new Date(isoDate);
  const day = date.getDay() || 7;
  const monday = new Date(date);
  monday.setDate(date.getDate() - day + 1);
  return monday.toISOString().slice(0, 10);
}

export function formatWeekLabel(weekKey: string) {
  const start = new Date(weekKey);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  const fmt = (d: Date) => d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  return `${fmt(start)} – ${fmt(end)}`;
}
