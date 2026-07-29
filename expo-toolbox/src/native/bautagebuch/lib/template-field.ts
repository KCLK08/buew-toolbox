import type {
  DetectedField,
  FieldGeometry,
  FieldRect,
  FieldSource,
  SetupFieldType
} from '../types';

export type TemplateFieldInput = {
  fieldId: string;
  fieldName: string;
  labelCandidate: string;
  type: string;
  options?: string[];
  page?: number;
  orderIndex?: number;
  source?: FieldSource;
  geometry?: FieldGeometry | null;
  /** @deprecated Legacy PDF rect [x1,y1,x2,y2] — normalized to geometry on save. */
  rect?: number[] | null;
};

export function createId(prefix = 'field'): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

export function legacyRectToFieldRect(rect: number[]): FieldRect {
  const [x1, y1, x2, y2] = rect.map((value) => Number(value));
  const x = Math.min(x1, x2);
  const y = Math.min(y1, y2);
  return {
    x,
    y,
    width: Math.abs(x2 - x1),
    height: Math.abs(y2 - y1)
  };
}

export function fieldRectToLegacyRect(rect: FieldRect): number[] {
  return [rect.x, rect.y, rect.x + rect.width, rect.y + rect.height];
}

export function normalizeFieldGeometry(input: {
  page?: number;
  geometry?: FieldGeometry | null;
  rect?: number[] | null;
}): FieldGeometry | null {
  if (input.geometry?.rect) {
    return {
      page: Math.max(1, Number(input.geometry.page || input.page || 1)),
      rect: {
        x: Number(input.geometry.rect.x),
        y: Number(input.geometry.rect.y),
        width: Math.max(0, Number(input.geometry.rect.width)),
        height: Math.max(0, Number(input.geometry.rect.height))
      }
    };
  }
  if (Array.isArray(input.rect) && input.rect.length >= 4) {
    return {
      page: Math.max(1, Number(input.page || 1)),
      rect: legacyRectToFieldRect(input.rect)
    };
  }
  return null;
}

type GeometryLike = {
  page?: number;
  geometry?: FieldGeometry | null;
  rect?: number[] | null;
};

export function getFieldGeometry(field: GeometryLike): FieldGeometry | null {
  return normalizeFieldGeometry({
    page: field.page,
    geometry: field.geometry ?? null,
    rect: field.rect ?? null
  });
}

export function getFieldPage(field: Pick<DetectedField, 'geometry' | 'page'>): number {
  return Math.max(1, Number(field.geometry?.page || field.page || 1));
}

export function fieldHasGeometry(field: GeometryLike): boolean {
  return getFieldGeometry(field) !== null;
}

export function fieldToPreviewLegacyRect(field: Pick<DetectedField, 'geometry' | 'page' | 'rect'>): number[] | null {
  const geometry = getFieldGeometry(field);
  if (!geometry) return null;
  return fieldRectToLegacyRect(geometry.rect);
}

export function normalizeDetectedField(row: Record<string, unknown>): DetectedField {
  const rectRaw = row.rectJson ? (JSON.parse(String(row.rectJson)) as number[]) : null;
  const geometryRaw = row.geometryJson
    ? (JSON.parse(String(row.geometryJson)) as FieldGeometry)
    : null;
  const geometry = normalizeFieldGeometry({
    page: Number(row.page || 1),
    geometry: geometryRaw,
    rect: rectRaw
  });
  const sourceRaw = String(row.source || '').trim() as FieldSource;
  let source: FieldSource = sourceRaw === 'manual' || sourceRaw === 'ocr' ? sourceRaw : 'acroform';
  if (!geometry && source === 'acroform' && !rectRaw) {
    source = 'acroform';
  }
  if (geometry && !sourceRaw) {
    source = 'acroform';
  }

  const legacyRect = geometry ? fieldRectToLegacyRect(geometry.rect) : rectRaw;

  return {
    id: String(row.id),
    templateId: String(row.templateId),
    fieldId: String(row.fieldId),
    fieldName: String(row.fieldName),
    labelCandidate: String(row.labelCandidate),
    type: String(row.type),
    options: JSON.parse(String(row.optionsJson || '[]')) as string[],
    page: getFieldPage({ geometry, page: Number(row.page || 1) }),
    orderIndex: Number(row.orderIndex || 0),
    rect: legacyRect,
    source,
    geometry,
    createdAt: String(row.createdAt),
    updatedAt: String(row.updatedAt)
  };
}

export function serializeFieldForDb(
  templateId: string,
  field: TemplateFieldInput,
  timestamp: string
): {
  id: string;
  templateId: string;
  fieldId: string;
  fieldName: string;
  labelCandidate: string;
  type: string;
  optionsJson: string;
  page: number;
  orderIndex: number;
  rectJson: string | null;
  geometryJson: string | null;
  source: FieldSource;
  createdAt: string;
  updatedAt: string;
} {
  const geometry = normalizeFieldGeometry({
    page: field.page,
    geometry: field.geometry,
    rect: field.rect
  });
  const resolvedSource: FieldSource =
    field.source || (geometry ? 'acroform' : 'acroform');
  const legacyRect = geometry ? fieldRectToLegacyRect(geometry.rect) : null;

  return {
    id: `${templateId}::${field.fieldId}`,
    templateId,
    fieldId: field.fieldId,
    fieldName: field.fieldName,
    labelCandidate: field.labelCandidate,
    type: field.type,
    optionsJson: JSON.stringify(field.options || []),
    page: geometry?.page || Math.max(1, Number(field.page || 1)),
    orderIndex: Number(field.orderIndex || 0),
    rectJson: legacyRect ? JSON.stringify(legacyRect) : null,
    geometryJson: geometry ? JSON.stringify(geometry) : null,
    source: resolvedSource,
    createdAt: timestamp,
    updatedAt: timestamp
  };
}

export function fieldSourceLabel(source: FieldSource): string {
  if (source === 'manual') return 'Manuell erstellt';
  if (source === 'ocr') return 'OCR';
  return 'Automatisch erkannt';
}

export function fieldSourceTone(source: FieldSource): 'success' | 'warning' | 'neutral' {
  if (source === 'manual') return 'warning';
  if (source === 'ocr') return 'neutral';
  return 'success';
}

export function countFieldsBySource(fields: DetectedField[]): {
  total: number;
  acroform: number;
  manual: number;
  ocr: number;
  withGeometry: number;
} {
  return {
    total: fields.length,
    acroform: fields.filter((field) => field.source === 'acroform').length,
    manual: fields.filter((field) => field.source === 'manual').length,
    ocr: fields.filter((field) => field.source === 'ocr').length,
    withGeometry: fields.filter((field) => fieldHasGeometry(field)).length
  };
}

export function scanResultToTemplateFieldInput(
  field: {
    fieldId: string;
    fieldName: string;
    labelCandidate: string;
    type: string;
    options: string[];
    page: number;
    orderIndex: number;
    rect: number[] | null;
  },
  source: FieldSource = 'acroform'
): TemplateFieldInput {
  const geometry = normalizeFieldGeometry({ page: field.page, rect: field.rect });
  return {
    fieldId: field.fieldId,
    fieldName: field.fieldName,
    labelCandidate: field.labelCandidate,
    type: field.type,
    options: field.options,
    page: field.page,
    orderIndex: field.orderIndex,
    source: geometry ? source : 'acroform',
    geometry
  };
}

export function createManualFieldInput(input: {
  name: string;
  type: SetupFieldType | string;
  page: number;
  rect: FieldRect;
  orderIndex?: number;
}): TemplateFieldInput {
  const fieldId = createId('manual');
  const name = String(input.name || '').trim() || 'Neues Feld';
  return {
    fieldId,
    fieldName: fieldId,
    labelCandidate: name,
    type: String(input.type || 'text'),
    options: [],
    page: Math.max(1, input.page),
    orderIndex: Number(input.orderIndex || 0),
    source: 'manual',
    geometry: {
      page: Math.max(1, input.page),
      rect: input.rect
    }
  };
}
