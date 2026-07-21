import * as SQLite from 'expo-sqlite';

import { nowIso } from '../../../lib/ids';

const DB_NAME = 'sitereport_native.db';

export type SiteReportColumn = {
  id: string;
  name: string;
  type: 'text' | 'number';
  isPhoto: boolean;
};

export type SiteReportEntry = {
  id: string;
  createdAt: string;
  fields: Record<string, string | number>;
  photoPath: string | null;
};

export type SiteReportProtocol = {
  id: string;
  createdAt: string;
  updatedAt: string;
  protocolTitle: string;
  projectName: string;
  protocolDate: string;
  protocolDescription: string;
  attendees: string;
  columns: SiteReportColumn[];
  entries: SiteReportEntry[];
  deleted_at: string | null;
};

let dbPromise: Promise<SQLite.SQLiteDatabase> | null = null;

function createId(prefix: string) {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

const defaultColumns: SiteReportColumn[] = [
  { id: 'col_photo', name: 'Bilder', type: 'text', isPhoto: true },
  { id: 'col_km', name: 'Kilometer', type: 'number', isPhoto: false },
  { id: 'col_desc', name: 'Beschreibung', type: 'text', isPhoto: false },
  { id: 'col_status', name: 'Status', type: 'text', isPhoto: false }
];

async function getDb() {
  if (!dbPromise) {
    dbPromise = SQLite.openDatabaseAsync(DB_NAME).then(async (db) => {
      await db.execAsync(`
        CREATE TABLE IF NOT EXISTS protocols (
          id TEXT PRIMARY KEY NOT NULL,
          createdAt TEXT NOT NULL,
          updatedAt TEXT NOT NULL,
          protocolTitle TEXT NOT NULL,
          projectName TEXT NOT NULL,
          protocolDate TEXT NOT NULL,
          protocolDescription TEXT NOT NULL DEFAULT '',
          attendees TEXT NOT NULL DEFAULT '',
          columnsJson TEXT NOT NULL,
          entriesJson TEXT NOT NULL,
          deleted_at TEXT
        );
      `);
      return db;
    });
  }
  return dbPromise;
}

function rowToProtocol(row: Record<string, unknown>): SiteReportProtocol {
  return {
    id: String(row.id),
    createdAt: String(row.createdAt),
    updatedAt: String(row.updatedAt),
    protocolTitle: String(row.protocolTitle),
    projectName: String(row.projectName),
    protocolDate: String(row.protocolDate),
    protocolDescription: String(row.protocolDescription || ''),
    attendees: String(row.attendees || ''),
    columns: JSON.parse(String(row.columnsJson || '[]')) as SiteReportColumn[],
    entries: JSON.parse(String(row.entriesJson || '[]')) as SiteReportEntry[],
    deleted_at: row.deleted_at ? String(row.deleted_at) : null
  };
}

export async function initSiteReportDatabase() {
  await getDb();
}

export async function listProtocols(): Promise<SiteReportProtocol[]> {
  const db = await getDb();
  const rows = await db.getAllAsync<Record<string, unknown>>(
    'SELECT * FROM protocols ORDER BY updatedAt DESC'
  );
  return rows.map(rowToProtocol).filter((row) => !row.deleted_at);
}

export async function getProtocol(id: string): Promise<SiteReportProtocol | null> {
  const db = await getDb();
  const row = await db.getFirstAsync<Record<string, unknown>>('SELECT * FROM protocols WHERE id = ?', id);
  if (!row) return null;
  const protocol = rowToProtocol(row);
  return protocol.deleted_at ? null : protocol;
}

export async function createProtocol(input: {
  protocolTitle: string;
  projectName: string;
  protocolDate: string;
  protocolDescription?: string;
  attendees?: string;
}): Promise<SiteReportProtocol> {
  const db = await getDb();
  const timestamp = nowIso();
  const protocol: SiteReportProtocol = {
    id: createId('protocol'),
    createdAt: timestamp,
    updatedAt: timestamp,
    protocolTitle: input.protocolTitle.trim(),
    projectName: input.projectName.trim(),
    protocolDate: input.protocolDate,
    protocolDescription: input.protocolDescription?.trim() || '',
    attendees: input.attendees?.trim() || '',
    columns: defaultColumns,
    entries: [],
    deleted_at: null
  };
  await db.runAsync(
    `INSERT INTO protocols (
      id, createdAt, updatedAt, protocolTitle, projectName, protocolDate,
      protocolDescription, attendees, columnsJson, entriesJson, deleted_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL)`,
    protocol.id,
    protocol.createdAt,
    protocol.updatedAt,
    protocol.protocolTitle,
    protocol.projectName,
    protocol.protocolDate,
    protocol.protocolDescription,
    protocol.attendees,
    JSON.stringify(protocol.columns),
    JSON.stringify(protocol.entries)
  );
  return protocol;
}

export async function updateProtocol(protocol: SiteReportProtocol): Promise<void> {
  const db = await getDb();
  await db.runAsync(
    `UPDATE protocols SET
      updatedAt = ?, protocolTitle = ?, projectName = ?, protocolDate = ?,
      protocolDescription = ?, attendees = ?, columnsJson = ?, entriesJson = ?
     WHERE id = ?`,
    nowIso(),
    protocol.protocolTitle,
    protocol.projectName,
    protocol.protocolDate,
    protocol.protocolDescription,
    protocol.attendees,
    JSON.stringify(protocol.columns),
    JSON.stringify(protocol.entries),
    protocol.id
  );
}

export async function softDeleteProtocol(id: string): Promise<void> {
  const db = await getDb();
  await db.runAsync('UPDATE protocols SET deleted_at = ?, updatedAt = ? WHERE id = ?', nowIso(), nowIso(), id);
}

export function todayDe(): string {
  const now = new Date();
  const day = String(now.getDate()).padStart(2, '0');
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const year = now.getFullYear();
  return `${day}-${month}-${year}`;
}
