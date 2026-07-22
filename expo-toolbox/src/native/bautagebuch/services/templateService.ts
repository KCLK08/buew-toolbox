import * as FileSystem from 'expo-file-system/legacy';

import {
  ETB_SETUP_VERSION,
  ETB_TEMPLATE_FILE_NAME,
  ETB_TEMPLATE_KIND,
  ETB_TEMPLATE_NAME,
  buildEtbSetupModel
} from '../lib/etb-template.js';
import { scanTemplatePdf } from '../lib/pdf-scan';
import { scanTemplatePdfLite } from '../lib/pdf-scan-lite';
import { detectedFieldsNeedRescan } from '../lib/scan-meta';
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
import { base64ToUint8Array, uint8ToBase64 } from '../../../lib/binary';

const TEMPLATE_URL = `${appConfig.toolboxWebBaseUrl.replace(/\/$/, '')}/bautagebuch/templates/Vorlage-eBTB.pdf`;

async function scanTemplateBytes(bytes: Uint8Array) {
  try {
    return await scanTemplatePdf(bytes);
  } catch {
    return scanTemplatePdfLite(bytes);
  }
}

async function readTemplateBytes(pdfPath: string): Promise<Uint8Array> {
  const base64 = await FileSystem.readAsStringAsync(pdfPath, {
    encoding: FileSystem.EncodingType.Base64
  });
  return base64ToUint8Array(base64);
}

async function rescanExistingTemplate(
  existing: NonNullable<Awaited<ReturnType<typeof listTemplates>>[number]>,
  setupModel: Record<string, unknown>
): Promise<{ templateId: string; setupModel: Record<string, unknown> }> {
  if (!existing.pdfPath) {
    throw new Error('Vorlage-eBTB ist lokal nicht verfügbar.');
  }

  const bytes = await readTemplateBytes(existing.pdfPath);
  const scanResult = await scanTemplateBytes(bytes);

  await putTemplate({
    ...existing,
    pageCount: scanResult.pageCount,
    updatedAt: new Date().toISOString()
  });

  await saveDetectedFields(
    existing.templateId,
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

  return { templateId: existing.templateId, setupModel };
}

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
    const detectedFields = await getDetectedFields(existing.templateId);
    const setupReady = setupModel && Number(setupModel.version || 0) >= ETB_SETUP_VERSION;
    const fieldsNeedRescan = detectedFieldsNeedRescan(detectedFields);

    if (setupReady && !fieldsNeedRescan) {
      return { templateId: existing.templateId, setupModel };
    }

    if (setupReady && fieldsNeedRescan && existing.pdfPath) {
      return rescanExistingTemplate(existing, setupModel);
    }
  }

  const response = await fetch(TEMPLATE_URL);
  if (!response.ok) {
    throw new Error('Vorlage-eBTB konnte nicht geladen werden.');
  }

  const arrayBuffer = await response.arrayBuffer();
  const bytes = new Uint8Array(arrayBuffer);
  const scanResult = await scanTemplateBytes(bytes);

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
