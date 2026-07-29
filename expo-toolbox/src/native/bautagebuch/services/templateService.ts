import * as DocumentPicker from 'expo-document-picker';
import * as FileSystem from 'expo-file-system/legacy';

import {
  ETB_SETUP_VERSION,
  ETB_TEMPLATE_FILE_NAME,
  ETB_TEMPLATE_KIND,
  ETB_TEMPLATE_NAME,
  buildEtbSetupModel,
  upgradeSetupModel
} from '../lib/etb-template.js';
import { buildGenericSetupModel } from '../lib/generic-setup-model.js';
import { mergeScannedFields } from '../lib/field-merge';
import { scanResultToTemplateFieldInput } from '../lib/template-field';
import { withWizardState } from '../lib/setup-mapping';
import { scanTemplatePdf } from '../lib/pdf-scan';
import { scanTemplatePdfLite } from '../lib/pdf-scan-lite';
import { detectedFieldsNeedRescan } from '../lib/scan-meta';
import {
  getAppSetting,
  getDetectedFields,
  getSetupModel,
  getTemplate,
  listTemplates,
  putTemplate,
  deleteTemplateRecord,
  renameTemplate as renameTemplateRecord,
  saveDetectedFields,
  saveSetupModel,
  setAppSetting
} from '../db/database';
import { appConfig } from '../../../lib/config';
import { base64ToUint8Array, uint8ToBase64 } from '../../../lib/binary';
import type { BautagebuchTemplate, BautagebuchTemplateStatus } from '../types';

const TEMPLATE_URL = `${appConfig.toolboxWebBaseUrl.replace(/\/$/, '')}/bautagebuch/templates/Vorlage-eBTB.pdf`;
const ACTIVE_TEMPLATE_KEY = 'activeTemplateId';
const MAX_TEMPLATE_BYTES = 40 * 1024 * 1024;

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

function sanitizeFileName(name: string): string {
  return String(name || 'vorlage.pdf')
    .trim()
    .replace(/[^\w.\-äöüÄÖÜß ]+/g, '_')
    .replace(/\s+/g, '_')
    .slice(0, 120);
}

function templateDisplayNameFromFile(fileName: string): string {
  return String(fileName || 'Vorlage')
    .replace(/\.pdf$/i, '')
    .replace(/[_-]+/g, ' ')
    .trim();
}

async function persistSetupModel(
  templateId: string,
  rawSetupModel: Record<string, unknown> | null,
  options: { upgradeEtb?: boolean } = {}
): Promise<Record<string, unknown>> {
  let setupModel = rawSetupModel;
  if (options.upgradeEtb && setupModel) {
    const upgraded = upgradeSetupModel(setupModel);
    setupModel = upgraded.model;
    if (upgraded.changed && setupModel) {
      await saveSetupModel(
        templateId,
        setupModel,
        String(setupModel.status || 'ready') as BautagebuchTemplateStatus
      );
    }
  }
  if (!setupModel) {
    throw new Error('Setup-Modell fehlt.');
  }
  return setupModel;
}

async function rescanExistingTemplate(
  existing: BautagebuchTemplate,
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

  const existingFields = await getDetectedFields(existing.templateId);
  const merged = mergeScannedFields(existingFields, scanResult.detectedFields);

  await saveDetectedFields(existing.templateId, merged);

  return { templateId: existing.templateId, setupModel };
}

export async function getActiveTemplateId(): Promise<string | null> {
  const value = await getAppSetting(ACTIVE_TEMPLATE_KEY);
  return value ? String(value).trim() : null;
}

export async function setActiveTemplateId(templateId: string): Promise<void> {
  const template = await getTemplate(templateId);
  if (!template) {
    throw new Error('Vorlage nicht gefunden.');
  }
  if (template.status !== 'ready') {
    throw new Error('Nur abgeschlossene Vorlagen können aktiviert werden.');
  }
  await setAppSetting(ACTIVE_TEMPLATE_KEY, template.templateId);
}

export async function archiveTemplate(templateId: string): Promise<void> {
  const template = await getTemplate(templateId);
  if (!template) {
    throw new Error('Vorlage nicht gefunden.');
  }
  const activeId = await getActiveTemplateId();
  if (activeId === templateId) {
    throw new Error('Die aktive Vorlage kann nicht archiviert werden.');
  }
  const setupModel = await getSetupModel(templateId);
  if (setupModel) {
    await saveSetupModel(templateId, setupModel, 'archived');
    return;
  }
  await putTemplate({
    ...template,
    status: 'archived',
    updatedAt: new Date().toISOString()
  });
}

function isBuiltinTemplate(template: BautagebuchTemplate): boolean {
  return template.templateKind === ETB_TEMPLATE_KIND || template.fileName === ETB_TEMPLATE_FILE_NAME;
}

export async function deleteTemplate(templateId: string): Promise<void> {
  const template = await getTemplate(templateId);
  if (!template) {
    throw new Error('Vorlage nicht gefunden.');
  }
  if (isBuiltinTemplate(template)) {
    throw new Error('Die Standard-Vorlage kann nicht gelöscht werden.');
  }
  const activeId = await getActiveTemplateId();
  if (activeId === templateId) {
    throw new Error('Die aktive Vorlage kann nicht gelöscht werden.');
  }

  await deleteTemplateRecord(templateId);

  if (template.pdfPath) {
    try {
      const info = await FileSystem.getInfoAsync(template.pdfPath);
      if (info.exists) {
        await FileSystem.deleteAsync(template.pdfPath, { idempotent: true });
      }
    } catch {
      // PDF cleanup is best-effort; DB soft-delete is authoritative.
    }
  }
}

export async function renameTemplate(
  templateId: string,
  templateName: string
): Promise<BautagebuchTemplate> {
  const updated = await renameTemplateRecord(templateId, templateName);
  if (!updated) {
    throw new Error('Vorlage nicht gefunden.');
  }
  return updated;
}

export async function listManagedTemplates(): Promise<BautagebuchTemplate[]> {
  return listTemplates();
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
    const rawSetupModel = await getSetupModel(existing.templateId);
    const setupModel = await persistSetupModel(existing.templateId, rawSetupModel, { upgradeEtb: true });
    const detectedFields = await getDetectedFields(existing.templateId);
    const setupReady = Number(setupModel.version || 0) >= ETB_SETUP_VERSION;
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
    scanResult.detectedFields.map((field) => scanResultToTemplateFieldInput(field, 'acroform'))
  );

  const detectedFields = await getDetectedFields(template.templateId);
  const setupModel = buildEtbSetupModel({
    templateId: template.templateId,
    pageCount: scanResult.pageCount,
    detectedFields: detectedFields as never[]
  });

  await saveSetupModel(template.templateId, setupModel, 'ready');

  const activeId = await getActiveTemplateId();
  if (!activeId) {
    await setActiveTemplateId(template.templateId);
  }

  return { templateId: template.templateId, setupModel };
}

export async function resolveActiveTemplateId(): Promise<string> {
  await ensureBuiltinTemplate();
  const activeId = await getActiveTemplateId();
  if (activeId) {
    const activeTemplate = await getTemplate(activeId);
    if (activeTemplate) {
      return activeTemplate.templateId;
    }
  }

  const builtin = (await listTemplates()).find(
    (template) =>
      template.templateKind === ETB_TEMPLATE_KIND || template.fileName === ETB_TEMPLATE_FILE_NAME
  );
  if (!builtin) {
    throw new Error('Keine Vorlage verfügbar.');
  }
  await setActiveTemplateId(builtin.templateId);
  return builtin.templateId;
}

export async function getTemplateBundle(templateId: string): Promise<{
  template: BautagebuchTemplate;
  setupModel: Record<string, unknown>;
}> {
  const template = await getTemplate(templateId);
  if (!template) {
    throw new Error('Vorlage nicht gefunden.');
  }

  const rawSetupModel = await getSetupModel(templateId);
  const upgradeEtb = template.templateKind === ETB_TEMPLATE_KIND;
  const setupModel = await persistSetupModel(templateId, rawSetupModel, { upgradeEtb });
  return { template, setupModel };
}

export async function getActiveTemplateBundle(): Promise<{
  template: BautagebuchTemplate;
  setupModel: Record<string, unknown>;
}> {
  const templateId = await resolveActiveTemplateId();
  return getTemplateBundle(templateId);
}

export async function importTemplateFromDocument(): Promise<{ templateId: string } | null> {
  const result = await DocumentPicker.getDocumentAsync({
    type: 'application/pdf',
    copyToCacheDirectory: true,
    multiple: false
  });

  if (result.canceled || !result.assets?.[0]) {
    return null;
  }

  const asset = result.assets[0];
  const uri = asset.uri;
  const fileName = sanitizeFileName(asset.name || 'vorlage.pdf');
  const mimeType = asset.mimeType || 'application/pdf';

  const base64 = await FileSystem.readAsStringAsync(uri, {
    encoding: FileSystem.EncodingType.Base64
  });
  const bytes = base64ToUint8Array(base64);
  if (bytes.byteLength > MAX_TEMPLATE_BYTES) {
    throw new Error('Die PDF-Vorlage überschreitet 40 MB und kann nicht importiert werden.');
  }

  const scanResult = await scanTemplateBytes(bytes);
  const pdfDir = `${FileSystem.documentDirectory}bautagebuch/templates/`;
  await FileSystem.makeDirectoryAsync(pdfDir, { intermediates: true });
  const templateId = `tplv2_${Date.now()}`;
  const pdfPath = `${pdfDir}${templateId}_${fileName}`;
  await FileSystem.writeAsStringAsync(pdfPath, uint8ToBase64(bytes), {
    encoding: FileSystem.EncodingType.Base64
  });

  const templateName = templateDisplayNameFromFile(fileName);
  const template = await putTemplate({
    templateId,
    templateName,
    fileName,
    templateKind: '',
    mimeType,
    sizeBytes: bytes.byteLength,
    pageCount: scanResult.pageCount,
    pdfPath,
    status: 'draft',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    deleted_at: null
  });

  await saveDetectedFields(
    template.templateId,
    scanResult.detectedFields.map((field) => scanResultToTemplateFieldInput(field, 'acroform'))
  );

  const detectedFields = await getDetectedFields(template.templateId);
  const setupModel = withWizardState(
    buildGenericSetupModel({
      templateId: template.templateId,
      templateName: template.templateName,
      pageCount: scanResult.pageCount,
      detectedFields: detectedFields as never[]
    }),
    { step: 'structure' }
  );

  await saveSetupModel(template.templateId, setupModel, 'in_progress');
  return { templateId: template.templateId };
}

export function isTemplateEditable(template: BautagebuchTemplate): boolean {
  return template.status !== 'archived';
}

export function canActivateTemplate(template: BautagebuchTemplate): boolean {
  return template.status === 'ready';
}

export function canDeleteTemplate(template: BautagebuchTemplate, activeTemplateId: string): boolean {
  if (template.templateId === activeTemplateId) return false;
  return !isBuiltinTemplate(template);
}
