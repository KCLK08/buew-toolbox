import * as FileSystem from 'expo-file-system/legacy';

import {
  ETB_SETUP_VERSION,
  ETB_TEMPLATE_FILE_NAME,
  ETB_TEMPLATE_KIND,
  ETB_TEMPLATE_NAME,
  buildEtbSetupModel
} from '../lib/etb-template.js';
import { scanTemplatePdfLite } from '../lib/pdf-scan-lite';
import {
  getDetectedFields,
  getSetupModel,
  getTemplate,
  listTemplates,
  putTemplate,
  saveDetectedFields,
  saveSetupModel
} from '../db/database';
import { appConfig } from '../../../lib/config';
import { uint8ToBase64 } from '../../../lib/binary';

const TEMPLATE_URL = `${appConfig.toolboxWebBaseUrl.replace(/\/$/, '')}/bautagebuch/templates/Vorlage-eBTB.pdf`;

export async function ensureBuiltinTemplate(): Promise<{
  templateId: string;
  setupModel: Record<string, unknown>;
}> {
  const existing = (await listTemplates()).find(
    (template) =>
      template.templateKind === ETB_TEMPLATE_KIND || template.fileName === ETB_TEMPLATE_FILE_NAME
  );

  if (existing) {
    const setupModel = await getSetupModel(existing.templateId);
    if (setupModel && Number(setupModel.version || 0) >= ETB_SETUP_VERSION) {
      return { templateId: existing.templateId, setupModel };
    }
  }

  const response = await fetch(TEMPLATE_URL);
  if (!response.ok) {
    throw new Error('Vorlage-eBTB konnte nicht geladen werden.');
  }

  const arrayBuffer = await response.arrayBuffer();
  const bytes = new Uint8Array(arrayBuffer);
  const scanResult = await scanTemplatePdfLite(bytes);

  const pdfDir = `${FileSystem.documentDirectory}bautagebuch/templates/`;
  await FileSystem.makeDirectoryAsync(pdfDir, { intermediates: true });
  const pdfPath = `${pdfDir}${ETB_TEMPLATE_FILE_NAME}`;
  await FileSystem.writeAsStringAsync(pdfPath, uint8ToBase64(bytes), {
    encoding: FileSystem.EncodingType.Base64
  });

  const templateId = existing?.templateId || `tplv2_${Date.now()}`;
  const template = await putTemplate({
    templateId,
    templateName: ETB_TEMPLATE_NAME,
    fileName: ETB_TEMPLATE_FILE_NAME,
    templateKind: ETB_TEMPLATE_KIND,
    mimeType: 'application/pdf',
    sizeBytes: bytes.byteLength,
    pageCount: scanResult.pageCount,
    pdfPath,
    status: 'ready',
    createdAt: existing?.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    deleted_at: null
  });

  await saveDetectedFields(
    template.templateId,
    scanResult.detectedFields.map((field) => ({
      fieldId: field.fieldId,
      fieldName: field.fieldName,
      labelCandidate: field.labelCandidate,
      type: field.type,
      options: field.options,
      page: field.page,
      orderIndex: field.orderIndex,
      rect: field.rect
    }))
  );

  const detectedFields = await getDetectedFields(template.templateId);
  const setupModel = buildEtbSetupModel({
    templateId: template.templateId,
    pageCount: scanResult.pageCount,
    detectedFields: detectedFields as never[]
  });

  await saveSetupModel(template.templateId, setupModel, 'ready');
  return { templateId: template.templateId, setupModel };
}

export async function getActiveTemplateBundle() {
  const { templateId, setupModel } = await ensureBuiltinTemplate();
  const template = await getTemplate(templateId);
  if (!template) {
    throw new Error('Bautagebuch-Vorlage fehlt.');
  }
  return { template, setupModel };
}
