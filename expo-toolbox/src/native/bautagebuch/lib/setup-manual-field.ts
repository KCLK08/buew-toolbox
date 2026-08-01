import type { FieldGeometry, FieldRect } from '../types';
import type { MappingField } from './setup-mapping';

export function mappingFieldGeometry(field: MappingField | null | undefined): FieldGeometry | null {
  if (!field) return null;
  if (field.geometry?.rect) {
    return {
      page: Math.max(1, Number(field.geometry.page || field.page || 1)),
      rect: field.geometry.rect
    };
  }
  if (field.rect && field.rect.length >= 4) {
    const [x1, y1, x2, y2] = field.rect;
    return {
      page: Math.max(1, Number(field.page || 1)),
      rect: {
        x: Math.min(x1, x2),
        y: Math.min(y1, y2),
        width: Math.abs(x2 - x1),
        height: Math.abs(y2 - y1)
      }
    };
  }
  return null;
}

export function geometryDraftFromField(field: MappingField): { page: number; rect: FieldRect } | null {
  const geometry = mappingFieldGeometry(field);
  if (!geometry) return null;
  return { page: geometry.page, rect: geometry.rect };
}

export function isManualMappingField(field: MappingField | null | undefined): boolean {
  return field?.source === 'manual';
}
