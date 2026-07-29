import * as FileSystem from 'expo-file-system/legacy';
import * as SQLite from 'expo-sqlite';

import type { BautagebuchExport, BautagebuchRun, BautagebuchTemplate, DetectedField, SetupModelRecord } from '../types';
import { nowIso } from '../../../lib/ids';
import { requestDatabaseBackup } from '../../../storage/backupService';
import {
  normalizeDetectedField,
  serializeFieldForDb,
  type TemplateFieldInput
} from '../lib/template-field';

const DB_NAME = 'bautagebuch_v2_native.db';

let dbPromise: Promise<SQLite.SQLiteDatabase> | null = null;

function createId(prefix = 'id'): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

function isActive(record: { deleted_at?: string | null } | null | undefined): boolean {
  return !String(record?.deleted_at || '').trim();
}

async function getDb(): Promise<SQLite.SQLiteDatabase> {
  if (!dbPromise) {
    dbPromise = (async () => {
      const db = await SQLite.openDatabaseAsync(DB_NAME);
      await db.execAsync('PRAGMA foreign_keys = ON;');
      await db.execAsync('PRAGMA journal_mode = WAL;');
      await db.execAsync(`
        CREATE TABLE IF NOT EXISTS templates (
          templateId TEXT PRIMARY KEY NOT NULL,
          templateName TEXT NOT NULL,
          fileName TEXT NOT NULL,
          templateKind TEXT NOT NULL DEFAULT '',
          mimeType TEXT NOT NULL,
          sizeBytes INTEGER NOT NULL DEFAULT 0,
          pageCount INTEGER NOT NULL DEFAULT 1,
          pdfPath TEXT NOT NULL,
          status TEXT NOT NULL DEFAULT 'draft',
          createdAt TEXT NOT NULL,
          updatedAt TEXT NOT NULL,
          deleted_at TEXT
        );
        CREATE TABLE IF NOT EXISTS detected_fields (
          id TEXT PRIMARY KEY NOT NULL,
          templateId TEXT NOT NULL,
          fieldId TEXT NOT NULL,
          fieldName TEXT NOT NULL,
          labelCandidate TEXT NOT NULL,
          type TEXT NOT NULL,
          optionsJson TEXT NOT NULL DEFAULT '[]',
          page INTEGER NOT NULL DEFAULT 1,
          orderIndex INTEGER NOT NULL DEFAULT 0,
          rectJson TEXT,
          createdAt TEXT NOT NULL,
          updatedAt TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS setup_models (
          templateId TEXT PRIMARY KEY NOT NULL,
          status TEXT NOT NULL DEFAULT 'draft',
          version INTEGER NOT NULL DEFAULT 1,
          setupModelJson TEXT NOT NULL,
          createdAt TEXT NOT NULL,
          updatedAt TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS runs (
          runId TEXT PRIMARY KEY NOT NULL,
          templateId TEXT NOT NULL,
          title TEXT NOT NULL,
          setupVersion INTEGER NOT NULL DEFAULT 1,
          valuesJson TEXT NOT NULL DEFAULT '{}',
          sectionIndex INTEGER NOT NULL DEFAULT 0,
          status TEXT NOT NULL DEFAULT 'draft',
          photoDocJson TEXT NOT NULL DEFAULT '{"enabled":null,"entries":[],"updatedAt":""}',
          createdAt TEXT NOT NULL,
          updatedAt TEXT NOT NULL,
          completedAt TEXT NOT NULL DEFAULT '',
          deleted_at TEXT
        );
        CREATE TABLE IF NOT EXISTS exports (
          exportId TEXT PRIMARY KEY NOT NULL,
          runId TEXT NOT NULL,
          fileName TEXT NOT NULL,
          filePath TEXT NOT NULL,
          exportedAt TEXT NOT NULL,
          deleted_at TEXT
        );
        CREATE TABLE IF NOT EXISTS photo_assets (
          id TEXT PRIMARY KEY NOT NULL,
          runId TEXT NOT NULL,
          entryId TEXT NOT NULL,
          mimeType TEXT NOT NULL,
          localPath TEXT NOT NULL,
          sizeBytes INTEGER NOT NULL DEFAULT 0,
          status TEXT NOT NULL DEFAULT 'ready',
          createdAt TEXT NOT NULL,
          updatedAt TEXT NOT NULL,
          deleted_at TEXT
        );
        CREATE TABLE IF NOT EXISTS app_settings (
          key TEXT PRIMARY KEY NOT NULL,
          value TEXT NOT NULL,
          updatedAt TEXT NOT NULL
        );
      `);
      await FileSystem.makeDirectoryAsync(`${FileSystem.documentDirectory || ''}bautagebuch/`, {
        intermediates: true
      }).catch(() => undefined);
      await migrateDetectedFieldsSchema(db);
      return db;
    })().catch((error) => {
      dbPromise = null;
      throw error;
    });
  }
  return dbPromise;
}

async function migrateDetectedFieldsSchema(db: SQLite.SQLiteDatabase): Promise<void> {
  const migrations = [
    `ALTER TABLE detected_fields ADD COLUMN source TEXT NOT NULL DEFAULT 'acroform'`,
    `ALTER TABLE detected_fields ADD COLUMN geometryJson TEXT`
  ];
  for (const statement of migrations) {
    try {
      await db.execAsync(statement);
    } catch {
      // Column already exists.
    }
  }

  const rows = await db.getAllAsync<Record<string, unknown>>(
    'SELECT id, page, rectJson, geometryJson, source FROM detected_fields WHERE geometryJson IS NULL AND rectJson IS NOT NULL'
  );
  for (const row of rows) {
    try {
      const rect = JSON.parse(String(row.rectJson)) as number[];
      if (!Array.isArray(rect) || rect.length < 4) continue;
      const [x1, y1, x2, y2] = rect;
      const geometry = {
        page: Math.max(1, Number(row.page || 1)),
        rect: {
          x: Math.min(x1, x2),
          y: Math.min(y1, y2),
          width: Math.abs(x2 - x1),
          height: Math.abs(y2 - y1)
        }
      };
      await db.runAsync(
        'UPDATE detected_fields SET geometryJson = ?, source = COALESCE(source, ?) WHERE id = ?',
        JSON.stringify(geometry),
        'acroform',
        String(row.id)
      );
    } catch {
      // Best-effort migration.
    }
  }
}

function rowToTemplate(row: Record<string, unknown>): BautagebuchTemplate {
  return {
    templateId: String(row.templateId),
    templateName: String(row.templateName),
    fileName: String(row.fileName),
    templateKind: String(row.templateKind || ''),
    mimeType: String(row.mimeType),
    sizeBytes: Number(row.sizeBytes || 0),
    pageCount: Number(row.pageCount || 1),
    pdfPath: String(row.pdfPath),
    status: String(row.status) as BautagebuchTemplate['status'],
    createdAt: String(row.createdAt),
    updatedAt: String(row.updatedAt),
    deleted_at: row.deleted_at ? String(row.deleted_at) : null
  };
}

function rowToRun(row: Record<string, unknown>): BautagebuchRun {
  return {
    runId: String(row.runId),
    templateId: String(row.templateId),
    title: String(row.title),
    setupVersion: Number(row.setupVersion || 1),
    values: JSON.parse(String(row.valuesJson || '{}')) as Record<string, unknown>,
    sectionIndex: Number(row.sectionIndex || 0),
    status: String(row.status) as BautagebuchRun['status'],
    photoDoc: JSON.parse(String(row.photoDocJson || '{}')) as BautagebuchRun['photoDoc'],
    createdAt: String(row.createdAt),
    updatedAt: String(row.updatedAt),
    completedAt: String(row.completedAt || ''),
    deleted_at: row.deleted_at ? String(row.deleted_at) : null
  };
}

export async function initBautagebuchDatabase(): Promise<void> {
  await getDb();
}

export async function resetBautagebuchDatabaseConnection(): Promise<void> {
  if (!dbPromise) return;
  const db = await dbPromise;
  await db.closeAsync();
  dbPromise = null;
}

export async function listTemplates(): Promise<BautagebuchTemplate[]> {
  const db = await getDb();
  const rows = await db.getAllAsync<Record<string, unknown>>(
    'SELECT * FROM templates ORDER BY updatedAt DESC'
  );
  return rows.map(rowToTemplate).filter(isActive);
}

export async function getAppSetting(key: string): Promise<string | null> {
  const db = await getDb();
  const row = await db.getFirstAsync<Record<string, unknown>>(
    'SELECT value FROM app_settings WHERE key = ?',
    key
  );
  return row?.value != null ? String(row.value) : null;
}

export async function setAppSetting(key: string, value: string): Promise<void> {
  const db = await getDb();
  const timestamp = nowIso();
  await db.runAsync(
    `INSERT INTO app_settings (key, value, updatedAt)
     VALUES (?, ?, ?)
     ON CONFLICT(key) DO UPDATE SET value = excluded.value, updatedAt = excluded.updatedAt`,
    key,
    value,
    timestamp
  );
}

export async function getTemplate(templateId: string): Promise<BautagebuchTemplate | null> {
  const db = await getDb();
  const row = await db.getFirstAsync<Record<string, unknown>>(
    'SELECT * FROM templates WHERE templateId = ?',
    templateId
  );
  if (!row || !isActive(rowToTemplate(row))) return null;
  return rowToTemplate(row);
}

export async function putTemplate(template: BautagebuchTemplate): Promise<BautagebuchTemplate> {
  const db = await getDb();
  const timestamp = nowIso();
  const record = {
    ...template,
    updatedAt: timestamp,
    createdAt: template.createdAt || timestamp
  };
  await db.runAsync(
    `INSERT INTO templates (
      templateId, templateName, fileName, templateKind, mimeType, sizeBytes, pageCount,
      pdfPath, status, createdAt, updatedAt, deleted_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(templateId) DO UPDATE SET
      templateName=excluded.templateName,
      fileName=excluded.fileName,
      templateKind=excluded.templateKind,
      mimeType=excluded.mimeType,
      sizeBytes=excluded.sizeBytes,
      pageCount=excluded.pageCount,
      pdfPath=excluded.pdfPath,
      status=excluded.status,
      updatedAt=excluded.updatedAt,
      deleted_at=excluded.deleted_at`,
    record.templateId,
    record.templateName,
    record.fileName,
    record.templateKind,
    record.mimeType,
    record.sizeBytes,
    record.pageCount,
    record.pdfPath,
    record.status,
    record.createdAt,
    record.updatedAt,
    record.deleted_at
  );
  return record;
}

export async function renameTemplate(
  templateId: string,
  templateName: string
): Promise<BautagebuchTemplate | null> {
  const template = await getTemplate(templateId);
  if (!template) return null;
  const trimmed = String(templateName || '').trim();
  if (!trimmed) {
    throw new Error('Vorlagenname fehlt.');
  }
  return putTemplate({ ...template, templateName: trimmed });
}

export async function deleteTemplateRecord(templateId: string): Promise<void> {
  const template = await getTemplate(templateId);
  if (!template) {
    throw new Error('Vorlage nicht gefunden.');
  }
  const db = await getDb();
  const timestamp = nowIso();
  await db.runAsync('UPDATE templates SET deleted_at = ?, updatedAt = ? WHERE templateId = ?', timestamp, timestamp, templateId);
}

export async function saveDetectedFields(
  templateId: string,
  fields: TemplateFieldInput[]
): Promise<void> {
  const db = await getDb();
  const timestamp = nowIso();
  await db.runAsync('DELETE FROM detected_fields WHERE templateId = ?', templateId);
  for (const [index, field] of fields.entries()) {
    const serialized = serializeFieldForDb(
      templateId,
      { ...field, orderIndex: field.orderIndex ?? index },
      timestamp
    );
    await db.runAsync(
      `INSERT INTO detected_fields (
        id, templateId, fieldId, fieldName, labelCandidate, type, optionsJson,
        page, orderIndex, rectJson, geometryJson, source, createdAt, updatedAt
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      serialized.id,
      serialized.templateId,
      serialized.fieldId,
      serialized.fieldName,
      serialized.labelCandidate,
      serialized.type,
      serialized.optionsJson,
      serialized.page,
      serialized.orderIndex,
      serialized.rectJson,
      serialized.geometryJson,
      serialized.source,
      serialized.createdAt,
      serialized.updatedAt
    );
  }
  void requestDatabaseBackup('manual');
}

export async function addTemplateField(
  templateId: string,
  field: TemplateFieldInput
): Promise<DetectedField> {
  const db = await getDb();
  const timestamp = nowIso();
  const rows = await db.getAllAsync<{ orderIndex: number }>(
    'SELECT orderIndex FROM detected_fields WHERE templateId = ? ORDER BY orderIndex DESC LIMIT 1',
    templateId
  );
  const nextOrder = Number(rows[0]?.orderIndex ?? -1) + 1;
  const serialized = serializeFieldForDb(
    templateId,
    { ...field, orderIndex: field.orderIndex ?? nextOrder },
    timestamp
  );
  await db.runAsync(
    `INSERT INTO detected_fields (
      id, templateId, fieldId, fieldName, labelCandidate, type, optionsJson,
      page, orderIndex, rectJson, geometryJson, source, createdAt, updatedAt
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    serialized.id,
    serialized.templateId,
    serialized.fieldId,
    serialized.fieldName,
    serialized.labelCandidate,
    serialized.type,
    serialized.optionsJson,
    serialized.page,
    serialized.orderIndex,
    serialized.rectJson,
    serialized.geometryJson,
    serialized.source,
    serialized.createdAt,
    serialized.updatedAt
  );
  void requestDatabaseBackup('manual');
  return normalizeDetectedField({
    ...serialized,
    optionsJson: serialized.optionsJson
  });
}

export async function updateTemplateField(
  templateId: string,
  fieldId: string,
  patch: Partial<Omit<TemplateFieldInput, 'fieldId' | 'source'>>
): Promise<DetectedField | null> {
  const db = await getDb();
  const row = await db.getFirstAsync<Record<string, unknown>>(
    'SELECT * FROM detected_fields WHERE templateId = ? AND fieldId = ?',
    templateId,
    fieldId
  );
  if (!row) return null;

  const existing = normalizeDetectedField(row);
  const timestamp = nowIso();
  const merged: TemplateFieldInput = {
    fieldId: existing.fieldId,
    fieldName: patch.fieldName ?? existing.fieldName,
    labelCandidate: patch.labelCandidate ?? existing.labelCandidate,
    type: patch.type ?? existing.type,
    options: patch.options ?? existing.options,
    page: patch.page ?? existing.page,
    orderIndex: patch.orderIndex ?? existing.orderIndex,
    source: existing.source,
    geometry: patch.geometry !== undefined ? patch.geometry : existing.geometry,
    rect: patch.rect !== undefined ? patch.rect : existing.rect
  };
  const serialized = serializeFieldForDb(templateId, merged, timestamp);
  await db.runAsync(
    `UPDATE detected_fields SET
      fieldName = ?, labelCandidate = ?, type = ?, optionsJson = ?,
      page = ?, orderIndex = ?, rectJson = ?, geometryJson = ?, updatedAt = ?
     WHERE templateId = ? AND fieldId = ?`,
    serialized.fieldName,
    serialized.labelCandidate,
    serialized.type,
    serialized.optionsJson,
    serialized.page,
    serialized.orderIndex,
    serialized.rectJson,
    serialized.geometryJson,
    timestamp,
    templateId,
    fieldId
  );
  void requestDatabaseBackup('manual');
  return normalizeDetectedField({
    ...serialized,
    optionsJson: serialized.optionsJson,
    createdAt: existing.createdAt
  });
}

export async function deleteTemplateField(templateId: string, fieldId: string): Promise<boolean> {
  const db = await getDb();
  const result = await db.runAsync(
    'DELETE FROM detected_fields WHERE templateId = ? AND fieldId = ?',
    templateId,
    fieldId
  );
  void requestDatabaseBackup('record_deleted');
  return Number(result.changes || 0) > 0;
}

export async function getDetectedFields(templateId: string): Promise<DetectedField[]> {
  const db = await getDb();
  const rows = await db.getAllAsync<Record<string, unknown>>(
    'SELECT * FROM detected_fields WHERE templateId = ? ORDER BY page ASC, orderIndex ASC',
    templateId
  );
  return rows.map((row) => normalizeDetectedField(row));
}

export async function saveSetupModel(
  templateId: string,
  setupModel: Record<string, unknown>,
  status: 'draft' | 'in_progress' | 'ready' | 'archived' = 'draft'
): Promise<void> {
  const db = await getDb();
  const timestamp = nowIso();
  await db.runAsync(
    `INSERT INTO setup_models (templateId, status, version, setupModelJson, createdAt, updatedAt)
     VALUES (?, ?, ?, ?, ?, ?)
     ON CONFLICT(templateId) DO UPDATE SET
       status=excluded.status,
       version=excluded.version,
       setupModelJson=excluded.setupModelJson,
       updatedAt=excluded.updatedAt`,
    templateId,
    status,
    Number(setupModel.version || 1),
    JSON.stringify(setupModel),
    String(setupModel.createdAt || timestamp),
    timestamp
  );
  await db.runAsync('UPDATE templates SET status = ?, updatedAt = ? WHERE templateId = ?', status, timestamp, templateId);
}

export async function getSetupModel(templateId: string): Promise<Record<string, unknown> | null> {
  const db = await getDb();
  const row = await db.getFirstAsync<Record<string, unknown>>(
    'SELECT setupModelJson FROM setup_models WHERE templateId = ?',
    templateId
  );
  if (!row) return null;
  return JSON.parse(String(row.setupModelJson || '{}')) as Record<string, unknown>;
}

export async function listRuns(templateId = ''): Promise<BautagebuchRun[]> {
  const db = await getDb();
  const rows = templateId
    ? await db.getAllAsync<Record<string, unknown>>(
        'SELECT * FROM runs WHERE templateId = ? ORDER BY updatedAt DESC',
        templateId
      )
    : await db.getAllAsync<Record<string, unknown>>('SELECT * FROM runs ORDER BY updatedAt DESC');
  return rows.map(rowToRun).filter(isActive);
}

export async function getRun(runId: string): Promise<BautagebuchRun | null> {
  const db = await getDb();
  const row = await db.getFirstAsync<Record<string, unknown>>('SELECT * FROM runs WHERE runId = ?', runId);
  if (!row || !isActive(rowToRun(row))) return null;
  return rowToRun(row);
}

export async function createRun(input: {
  templateId: string;
  title: string;
  setupVersion?: number;
}): Promise<BautagebuchRun> {
  const db = await getDb();
  const timestamp = nowIso();
  const run: BautagebuchRun = {
    runId: createId('runv2'),
    templateId: input.templateId,
    title: String(input.title || '').trim() || 'BTB',
    setupVersion: Number(input.setupVersion || 1),
    values: {},
    sectionIndex: 0,
    status: 'draft',
    photoDoc: { enabled: null, entries: [], updatedAt: timestamp },
    createdAt: timestamp,
    updatedAt: timestamp,
    completedAt: '',
    deleted_at: null
  };
  await db.runAsync(
    `INSERT INTO runs (
      runId, templateId, title, setupVersion, valuesJson, sectionIndex, status,
      photoDocJson, createdAt, updatedAt, completedAt, deleted_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    run.runId,
    run.templateId,
    run.title,
    run.setupVersion,
    JSON.stringify(run.values),
    run.sectionIndex,
    run.status,
    JSON.stringify(run.photoDoc),
    run.createdAt,
    run.updatedAt,
    run.completedAt,
    run.deleted_at
  );
  return run;
}

async function softDeletePhotoAsset(runId: string, entryId: string): Promise<void> {
  const db = await getDb();
  const timestamp = nowIso();
  const id = `${runId}::${entryId}`;
  await db.runAsync(
    `UPDATE photo_assets SET status = 'deleted', deleted_at = ?, updatedAt = ? WHERE id = ?`,
    timestamp,
    timestamp,
    id
  );
}

async function syncPhotoDocAssets(
  runId: string,
  previousEntries: Array<{ id: string }>,
  nextEntries: Array<{ id: string }>
): Promise<void> {
  const nextIds = new Set(nextEntries.map((entry) => entry.id));
  for (const entry of previousEntries) {
    if (!nextIds.has(entry.id)) {
      await softDeletePhotoAsset(runId, entry.id);
    }
  }
}

export async function deletePhotoAsset(runId: string, entryId: string, localPath?: string): Promise<void> {
  await softDeletePhotoAsset(runId, entryId);
  if (localPath) {
    await FileSystem.deleteAsync(localPath, { idempotent: true });
  }
}

export async function updateRun(
  runId: string,
  patch: Partial<Pick<BautagebuchRun, 'title' | 'values' | 'sectionIndex' | 'status' | 'photoDoc' | 'completedAt'>>
): Promise<BautagebuchRun | null> {
  const existing = await getRun(runId);
  if (!existing) return null;

  if (patch.photoDoc) {
    await syncPhotoDocAssets(
      runId,
      existing.photoDoc?.entries || [],
      patch.photoDoc.entries || []
    );
  }

  const next: BautagebuchRun = {
    ...existing,
    ...patch,
    updatedAt: nowIso()
  };
  const db = await getDb();
  await db.runAsync(
    `UPDATE runs SET
      title = ?, valuesJson = ?, sectionIndex = ?, status = ?, photoDocJson = ?,
      updatedAt = ?, completedAt = ?
     WHERE runId = ?`,
    next.title,
    JSON.stringify(next.values),
    next.sectionIndex,
    next.status,
    JSON.stringify(next.photoDoc),
    next.updatedAt,
    next.completedAt,
    runId
  );
  return next;
}

export async function softDeleteRun(runId: string): Promise<void> {
  const db = await getDb();
  const timestamp = nowIso();
  await db.runAsync(
    'UPDATE runs SET deleted_at = ?, updatedAt = ?, status = ? WHERE runId = ?',
    timestamp,
    timestamp,
    'deleted',
    runId
  );
}

export async function deleteRunCascade(runId: string): Promise<void> {
  const run = await getRun(runId);
  if (!run) return;

  const db = await getDb();
  const timestamp = nowIso();

  for (const entry of run.photoDoc?.entries || []) {
    if (entry.localPath) {
      await FileSystem.deleteAsync(entry.localPath, { idempotent: true });
    }
    if (entry.id) {
      await softDeletePhotoAsset(runId, entry.id);
    }
  }

  await db.runAsync('UPDATE exports SET deleted_at = ? WHERE runId = ?', timestamp, runId);
  await softDeleteRun(runId);

  const photoDir = `${FileSystem.documentDirectory}bautagebuch/photos/${runId}/`;
  await FileSystem.deleteAsync(photoDir, { idempotent: true });
  void requestDatabaseBackup('record_deleted');
}

export async function renameRun(runId: string, title: string): Promise<BautagebuchRun | null> {
  const trimmed = String(title || '').trim();
  if (!trimmed) return null;
  return updateRun(runId, { title: trimmed });
}

export async function savePhotoAsset(input: {
  runId: string;
  entryId: string;
  localPath: string;
  mimeType: string;
  sizeBytes: number;
}): Promise<void> {
  const db = await getDb();
  const timestamp = nowIso();
  const id = `${input.runId}::${input.entryId}`;
  await db.runAsync(
    `INSERT INTO photo_assets (id, runId, entryId, mimeType, localPath, sizeBytes, status, createdAt, updatedAt, deleted_at)
     VALUES (?, ?, ?, ?, ?, ?, 'ready', ?, ?, NULL)
     ON CONFLICT(id) DO UPDATE SET
       localPath=excluded.localPath,
       mimeType=excluded.mimeType,
       sizeBytes=excluded.sizeBytes,
       status='ready',
       updatedAt=excluded.updatedAt,
       deleted_at=NULL`,
    id,
    input.runId,
    input.entryId,
    input.mimeType,
    input.localPath,
    input.sizeBytes,
    timestamp,
    timestamp
  );
}

export function getBautagebuchStorageRoot(): string {
  return `${FileSystem.documentDirectory}bautagebuch/`;
}

function rowToExport(row: Record<string, unknown>): BautagebuchExport {
  return {
    exportId: String(row.exportId),
    runId: String(row.runId),
    fileName: String(row.fileName),
    filePath: String(row.filePath),
    exportedAt: String(row.exportedAt),
    deleted_at: row.deleted_at ? String(row.deleted_at) : null
  };
}

export async function listExports(): Promise<BautagebuchExport[]> {
  const db = await getDb();
  const rows = await db.getAllAsync<Record<string, unknown>>(
    'SELECT * FROM exports WHERE deleted_at IS NULL ORDER BY exportedAt DESC'
  );
  return rows.map(rowToExport);
}

export async function getExport(exportId: string): Promise<BautagebuchExport | null> {
  const db = await getDb();
  const row = await db.getFirstAsync<Record<string, unknown>>(
    'SELECT * FROM exports WHERE exportId = ? AND deleted_at IS NULL',
    exportId
  );
  return row ? rowToExport(row) : null;
}

export async function getExportByRun(runId: string): Promise<BautagebuchExport | null> {
  const db = await getDb();
  const row = await db.getFirstAsync<Record<string, unknown>>(
    'SELECT * FROM exports WHERE runId = ? AND deleted_at IS NULL ORDER BY exportedAt DESC LIMIT 1',
    runId
  );
  return row ? rowToExport(row) : null;
}

export async function upsertExportByRun(record: Omit<BautagebuchExport, 'deleted_at'>): Promise<BautagebuchExport> {
  const db = await getDb();
  const existing = await getExportByRun(record.runId);
  const exportId = existing?.exportId || record.exportId || createId('export');
  const next: BautagebuchExport = {
    exportId,
    runId: record.runId,
    fileName: record.fileName,
    filePath: record.filePath,
    exportedAt: record.exportedAt,
    deleted_at: null
  };
  await db.runAsync(
    `INSERT OR REPLACE INTO exports (exportId, runId, fileName, filePath, exportedAt, deleted_at)
     VALUES (?, ?, ?, ?, ?, NULL)`,
    next.exportId,
    next.runId,
    next.fileName,
    next.filePath,
    next.exportedAt
  );
  return next;
}

export async function deleteExport(exportId: string): Promise<void> {
  const db = await getDb();
  await db.runAsync('UPDATE exports SET deleted_at = ? WHERE exportId = ?', nowIso(), exportId);
}
