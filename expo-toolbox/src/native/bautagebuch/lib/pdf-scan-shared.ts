import { PDFDropdown, PDFRadioGroup } from 'pdf-lib';

export type MutableScanField = {
  fieldName: string;
  labelCandidate: string;
  type: string;
  options: string[];
  page: number;
  orderIndex: number;
  rect: number[] | null;
  fieldId?: string;
};

export function humanizeFieldName(value: string): string {
  return String(value || '')
    .replace(/[_\.]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function slugify(value: string): string {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

export function uniqueStrings(values: string[] = []): string[] {
  return [...new Set(values.map((value) => String(value || '').trim()).filter(Boolean))];
}

export function normalizeRect(rect: number[] | null | undefined) {
  if (!Array.isArray(rect) || rect.length < 4) {
    return null;
  }
  const x1 = Number(rect[0]);
  const y1 = Number(rect[1]);
  const x2 = Number(rect[2]);
  const y2 = Number(rect[3]);
  if (![x1, y1, x2, y2].every(Number.isFinite)) {
    return null;
  }
  const left = Math.min(x1, x2);
  const right = Math.max(x1, x2);
  const top = Math.max(y1, y2);
  const bottom = Math.min(y1, y2);
  return {
    left,
    right,
    top,
    bottom,
    width: Math.max(0, right - left),
    height: Math.max(0, top - bottom),
    centerX: (left + right) / 2,
    centerY: (top + bottom) / 2
  };
}

export function readSelectOptions(
  field: { getOptions?: () => string[] },
  fallbackOptions: unknown[] = []
): string[] {
  try {
    if (field instanceof PDFDropdown || field instanceof PDFRadioGroup) {
      return uniqueStrings(field.getOptions());
    }
    if (typeof field.getOptions === 'function') {
      return uniqueStrings(field.getOptions());
    }
  } catch {
    // Some fields do not expose options.
  }

  if (Array.isArray(fallbackOptions)) {
    return uniqueStrings(
      fallbackOptions.map((option) => {
        if (typeof option === 'string') return option;
        const record = option as { displayValue?: string; exportValue?: string; value?: string };
        return record.displayValue || record.exportValue || record.value || '';
      })
    );
  }
  return [];
}

export function sortDetectedFields(fields: MutableScanField[]): MutableScanField[] {
  return [...fields].sort((left, right) => {
    if ((left.page ?? 9999) !== (right.page ?? 9999)) {
      return (left.page ?? 9999) - (right.page ?? 9999);
    }

    const leftRect = normalizeRect(left.rect);
    const rightRect = normalizeRect(right.rect);
    if (Boolean(leftRect) !== Boolean(rightRect)) {
      return leftRect ? -1 : 1;
    }

    if (leftRect && rightRect) {
      const topDelta = rightRect.top - leftRect.top;
      if (Math.abs(topDelta) > 4) return topDelta;
      const leftDelta = leftRect.left - rightRect.left;
      if (Math.abs(leftDelta) > 4) return leftDelta;
    }

    if ((left.orderIndex ?? 9999) !== (right.orderIndex ?? 9999)) {
      return (left.orderIndex ?? 9999) - (right.orderIndex ?? 9999);
    }

    return String(left.fieldName || '').localeCompare(String(right.fieldName || ''), 'de');
  });
}

export function assignFieldIds(fields: MutableScanField[]): Array<MutableScanField & { fieldId: string }> {
  const usedIds = new Set<string>();
  return fields.map((field, index) => {
    const slug = slugify(field.fieldName) || `field-${index + 1}`;
    const page = Number(field.page || 1);
    const orderIndex = Number(field.orderIndex ?? index + 1);
    const baseId = `${slug}-p${page}-o${orderIndex}`;
    let fieldId = baseId;
    let suffix = 2;
    while (usedIds.has(fieldId)) {
      fieldId = `${baseId}-${suffix}`;
      suffix += 1;
    }
    usedIds.add(fieldId);
    return { ...field, fieldId };
  });
}

export function inferLabelCandidate(
  textLines: Array<{ text: string; x: number; y: number; width: number; centerX: number }>,
  rect: number[] | null,
  fallback = ''
): string {
  const normalizedRect = normalizeRect(rect);
  if (!normalizedRect) {
    return String(fallback || '').trim();
  }

  const fallbackLabel = String(fallback || '').trim() || 'Feld';
  const candidates: Array<{ text: string; score: number }> = [];

  for (const line of textLines) {
    const text = String(line.text || '').trim();
    if (!text || text.length < 2) continue;

    const left = Number(line.x ?? 0);
    const right = Number(line.x ?? 0) + Number(line.width ?? 0);
    const centerX = Number(line.centerX ?? left);
    const y = Number(line.y ?? 0);
    const isLeftLabel =
      right <= normalizedRect.left + 12 &&
      Math.abs(y - normalizedRect.centerY) <= Math.max(10, normalizedRect.height * 1.15);
    const isTopLabel =
      y > normalizedRect.top + 1 &&
      y <= normalizedRect.top + 45 &&
      centerX >= normalizedRect.left - 140 &&
      centerX <= normalizedRect.right + 140;

    if (!isLeftLabel && !isTopLabel) continue;

    const score = isLeftLabel
      ? Math.abs(normalizedRect.left - right) + Math.abs(y - normalizedRect.centerY) * 1.4
      : 20 + Math.abs(y - normalizedRect.top) * 1.3 + Math.abs(centerX - normalizedRect.centerX) * 0.55;

    candidates.push({ text, score });
  }

  candidates.sort((left, right) => left.score - right.score);
  const best = candidates[0]?.text;
  if (!best || best.length > 80) return fallbackLabel;
  return best;
}
