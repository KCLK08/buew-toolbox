import type { DetectedField, SetupFieldConfig, SetupStructureItem } from '../types';
import {
  normalizeSetupFieldType,
  setupFieldTypeLabel
} from './setup-field-settings';
import { listSetupSections, resolveOverlayPlacement } from './setup-mapping';
import { getStructureItems } from './setup-structure';

export type TemplateDetailField = {
  fieldId: string;
  label: string;
  typeLabel: string;
  config: SetupFieldConfig;
};

export type TemplateDetailGroup = {
  kind: 'group';
  id: string;
  name: string;
  description?: string;
  fieldCount: number;
  fields: TemplateDetailField[];
};

export type TemplateDetailTable = {
  kind: 'table';
  id: string;
  name: string;
  columns: Array<{ id: string; name: string; type?: string }>;
};

export type TemplateDetailItem = TemplateDetailGroup | TemplateDetailTable;

export type TemplateDetailOverview = {
  items: TemplateDetailItem[];
  updatedAt: string | null;
};

const GROUP_ICONS = ['file-document-outline', 'weather-partly-cloudy', 'clipboard-text-outline', 'draw'];

export function structureItemIcon(item: SetupStructureItem, index: number): string {
  if (item.type === 'table') return 'table';
  const name = item.name.toLowerCase();
  if (name.includes('wetter')) return 'weather-partly-cloudy';
  if (name.includes('arbeit') || name.includes('leistung')) return 'clipboard-text-outline';
  if (name.includes('unterschrift') || name.includes('signatur')) return 'draw';
  return GROUP_ICONS[index % GROUP_ICONS.length];
}

export function structureItemSummary(item: SetupStructureItem, fieldCount: number): string {
  if (item.type === 'table') return 'Tabelle';
  return fieldCount === 1 ? '1 Feld' : `${fieldCount} Felder`;
}

export function resolveFieldPositionLabel(rect: number[] | null | undefined): string {
  if (!rect || rect.length < 4) return '—';
  const [x1, y1, x2, y2] = rect;
  const centerX = (x1 + x2) / 2;
  const centerY = (y1 + y2) / 2;
  const topThird = 842 * 0.68;
  const bottomThird = 842 * 0.32;
  const leftThird = 595 * 0.33;
  const rightThird = 595 * 0.67;

  const parts: string[] = [];
  if (centerY >= topThird) parts.push('oben');
  else if (centerY <= bottomThird) parts.push('unten');
  if (centerX <= leftThird) parts.push('links');
  else if (centerX >= rightThird) parts.push('rechts');
  return parts.length > 0 ? parts.join(' ') : 'Mitte';
}

export function resolveFieldOverlayHint(rect: number[] | null | undefined): string {
  const placement = resolveOverlayPlacement(rect || null);
  if (placement === 'top') return 'unten';
  if (placement === 'bottom') return 'oben';
  if (placement === 'left') return 'rechts';
  if (placement === 'right') return 'links';
  return 'Mitte';
}

export function formatTemplateUpdatedAt(iso: string | null | undefined): string {
  if (!iso) return '—';
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return '—';
  return date.toLocaleDateString('de-DE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  });
}

function fieldToDetail(field: SetupFieldConfig, detectedFields: DetectedField[]): TemplateDetailField {
  const type = normalizeSetupFieldType(field, detectedFields);
  return {
    fieldId: String(field.fieldId || ''),
    label: String(field.label || field.fieldName || 'Feld'),
    typeLabel: setupFieldTypeLabel(type),
    config: field
  };
}

function tableSectionById(setupModel: Record<string, unknown>, tableId: string) {
  const tables = Array.isArray(setupModel.table_sections) ? setupModel.table_sections : [];
  return tables.find((entry) => String(entry?.tableId || '') === tableId);
}

export function buildTemplateDetailOverview(
  setupModel: Record<string, unknown>,
  detectedFields: DetectedField[] = []
): TemplateDetailOverview {
  const structure = getStructureItems(setupModel);
  const sections = listSetupSections(setupModel);
  const sectionById = new Map(sections.map((section) => [section.sectionId, section]));

  const items: TemplateDetailItem[] = structure.map((item) => {
    if (item.type === 'group') {
      const section = sectionById.get(item.id);
      const fields = (section?.fields || [])
        .filter((field) => field.skipped !== true)
        .map((field) => fieldToDetail(field, detectedFields));
      return {
        kind: 'group',
        id: item.id,
        name: item.name,
        description: item.description,
        fieldCount: fields.length,
        fields
      };
    }

    const table = tableSectionById(setupModel, item.id);
    const columns = item.columns.map((column) => {
      const meta = (Array.isArray(table?.columns) ? table.columns : []).find(
        (entry: { columnId?: string }) => String(entry?.columnId) === column.id
      );
      return {
        id: column.id,
        name: column.name,
        type: meta?.type ? String(meta.type) : 'text'
      };
    });

    return {
      kind: 'table',
      id: item.id,
      name: item.name,
      columns
    };
  });

  return {
    items,
    updatedAt: setupModel.updatedAt ? String(setupModel.updatedAt) : null
  };
}

export type FieldDetailRow = { label: string; value: string; checked?: boolean };

export function buildFieldDetailRows(
  field: SetupFieldConfig,
  detectedFields: DetectedField[] = []
): FieldDetailRow[] {
  const type = normalizeSetupFieldType(field, detectedFields);
  const detected = detectedFields.find((entry) => entry.fieldId === field.fieldId);
  const page = Number(field.page || detected?.page || 1);

  const rows: FieldDetailRow[] = [
    { label: 'Quelle', value: `PDF Seite ${page}` },
    { label: 'Position', value: resolveFieldPositionLabel(field.rect || detected?.rect) },
    { label: 'Feldtyp', value: setupFieldTypeLabel(type) }
  ];

  const settings: FieldDetailRow[] = [
    { label: 'Pflichtfeld', value: '', checked: field.required === true },
    {
      label: 'Im BTB',
      value: field.skipped === true ? 'Ausgeblendet' : 'Sichtbar'
    }
  ];

  if (type === 'datetime') {
    settings.push({
      label: 'Heutiges Datum übernehmen',
      value: '',
      checked: field.useCurrentDate === true
    });
    const mode = field.dateMode || 'date';
    settings.push({
      label: 'Anzeige',
      value: mode === 'time' ? 'Zeit' : mode === 'datetime' ? 'Datum/Zeit' : 'Datum'
    });
  }

  if (type === 'select' && field.options?.length) {
    settings.push({ label: 'Optionen', value: field.options.join(', ') });
  }

  if (type === 'static_text' && field.staticText) {
    settings.push({ label: 'Statischer Text', value: field.staticText });
  }

  if (type === 'signature') {
    settings.push({
      label: 'Modus',
      value: field.signatureMode === 'image' ? 'Bild' : 'Zeichnen'
    });
  }

  if (field.multiline) {
    settings.push({ label: 'Mehrzeilig', value: '', checked: true });
  }

  if (field.defaultValue) {
    settings.push({ label: 'Standardwert', value: field.defaultValue });
  }

  if (field.hint) {
    settings.push({ label: 'Hinweis', value: field.hint });
  }

  return [...rows, ...settings];
}

export function buildGroupFieldSummaryRows(field: TemplateDetailField): FieldDetailRow[] {
  const config = field.config;
  return [
    { label: 'Typ', value: field.typeLabel },
    { label: 'Pflichtfeld', value: config.required === true ? 'Ja' : 'Nein' },
    {
      label: 'Im BTB',
      value: config.skipped === true ? 'Ausgeblendet' : 'Sichtbar'
    },
    ...(config.useCurrentDate
      ? [{ label: 'Automatisch', value: 'Heutiges Datum' }]
      : [])
  ];
}
