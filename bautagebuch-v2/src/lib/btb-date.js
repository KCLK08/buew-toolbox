import { parseBtbTitle } from './btb-naming.js';

function parseUtcDateParts(year, month, day) {
  const y = Number(year);
  const m = Number(month);
  const d = Number(day);
  if (!Number.isFinite(y) || !Number.isFinite(m) || !Number.isFinite(d)) return null;
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;

  const date = new Date(Date.UTC(y, m - 1, d));
  if (date.getUTCFullYear() !== y || date.getUTCMonth() !== m - 1 || date.getUTCDate() !== d) return null;
  return date;
}

export function parseBtbDateValue(value) {
  const raw = String(value || '').trim();
  if (!raw) return null;

  const dmy = raw.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})$/);
  if (dmy) {
    const day = Number(dmy[1]);
    const month = Number(dmy[2]);
    const yearToken = Number(dmy[3]);
    const year = dmy[3].length === 2 ? 2000 + yearToken : yearToken;
    return parseUtcDateParts(year, month, day);
  }

  const ymd = raw.match(/^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})$/);
  if (ymd) {
    return parseUtcDateParts(ymd[1], ymd[2], ymd[3]);
  }

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return null;
  return parseUtcDateParts(parsed.getFullYear(), parsed.getMonth() + 1, parsed.getDate());
}

export function btbDateFromRunValues(values = {}) {
  const entries = Object.entries(values || {}).filter(([, value]) => String(value ?? '').trim().length > 0);
  if (entries.length === 0) return null;

  const keyPatterns = [/^Date1$/i, /^field:date1(?:-|$)/i, /date1/i];
  for (const pattern of keyPatterns) {
    for (const [key, value] of entries) {
      if (!pattern.test(String(key || '').trim())) continue;
      const parsed = parseBtbDateValue(value);
      if (parsed) return parsed;
    }
  }
  return null;
}

/** @param {{ title?: string; createdAt?: string; values?: Record<string, unknown> }} run */
export function parseBtbDateFromRun(run) {
  const fromValues = btbDateFromRunValues(run?.values || {});
  if (fromValues) return fromValues;

  const parsedTitle = parseBtbTitle(String(run?.title || ''));
  if (parsedTitle?.date) {
    return parseUtcDateParts(
      parsedTitle.date.slice(0, 4),
      parsedTitle.date.slice(5, 7),
      parsedTitle.date.slice(8, 10)
    );
  }

  const fromTitle = String(run?.title || '').match(/\b(\d{4})-(\d{2})-(\d{2})\b/);
  if (fromTitle) {
    return parseUtcDateParts(fromTitle[1], fromTitle[2], fromTitle[3]);
  }

  const fallback = run?.createdAt?.slice(0, 10);
  if (fallback) {
    return parseUtcDateParts(fallback.slice(0, 4), fallback.slice(5, 7), fallback.slice(8, 10));
  }
  return null;
}
