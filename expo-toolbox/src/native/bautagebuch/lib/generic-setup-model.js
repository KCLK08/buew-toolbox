import { buildLegacySectionOrder } from './setup-model.js';

function createId(prefix = 'id') {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

function nowIso() {
  return new Date().toISOString();
}

function fieldEntry(sourceField, { label = '' } = {}) {
  if (sourceField) {
    return {
      fieldId: String(sourceField.fieldId || ''),
      fieldName: String(sourceField.fieldName || ''),
      label: String(label || sourceField.labelCandidate || sourceField.fieldName || 'Feld').trim(),
      type: String(sourceField.type || 'text'),
      options: Array.isArray(sourceField.options) ? [...sourceField.options] : [],
      required: false,
      skipped: false,
      multiline: false,
      page: Number(sourceField.page || 1),
      rect: Array.isArray(sourceField.rect) ? sourceField.rect.slice(0, 4) : null
    };
  }

  return {
    fieldId: '',
    fieldName: '',
    label: String(label || '').trim() || 'Feld',
    type: 'text',
    options: [],
    required: false,
    skipped: true,
    multiline: false,
    page: 1,
    rect: null
  };
}

export function buildGenericSetupModel({
  templateId,
  templateName,
  pageCount,
  detectedFields = []
}) {
  const fields = (Array.isArray(detectedFields) ? detectedFields : [])
    .filter((field) => String(field?.fieldId || '').trim())
    .filter((field) => String(field?.type || 'text') !== 'unsupported')
    .map((field) =>
      fieldEntry(field, {
        label: String(field.labelCandidate || field.fieldName || 'Feld').trim()
      })
    )
    .filter((field) => !field.skipped);

  const byPage = new Map();
  for (const field of fields) {
    const page = Number(field.page || 1);
    const bucket = byPage.get(page) || [];
    bucket.push(field);
    byPage.set(page, bucket);
  }

  const singleSections = [...byPage.entries()]
    .sort((left, right) => left[0] - right[0])
    .map(([page, pageFields]) => ({
      sectionId: `page_${page}`,
      label: `Seite ${page}`,
      page,
      fields: pageFields
    }));

  if (singleSections.length === 0) {
    singleSections.push({
      sectionId: 'fields',
      label: 'PDF-Felder',
      page: 1,
      fields: []
    });
  }

  return {
    modelId: createId('setupv2'),
    version: 1,
    status: 'draft',
    templateId: String(templateId || ''),
    templateName: String(templateName || 'Vorlage'),
    pageCount: Number(pageCount || 1),
    single_sections: singleSections,
    table_sections: [],
    section_order: buildLegacySectionOrder({ single_sections: singleSections, table_sections: [] }),
    createdAt: nowIso(),
    updatedAt: nowIso()
  };
}
