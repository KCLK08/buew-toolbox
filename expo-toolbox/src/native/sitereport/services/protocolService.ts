import {
  deleteExportsByProtocol,
  getProtocol,
  listExports,
  softDeleteProtocol,
  updateProtocol,
  type SiteReportEntry,
  type SiteReportProtocol
} from '../db/database';
import { deleteCachedExport } from './exportService';
import { deleteEntryPhoto, deleteProtocolPhotos } from './photoService';

export function createEntryId(): string {
  return `entry_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

export function emptyFieldsFromColumns(
  columns: SiteReportProtocol['columns']
): Record<string, string | number> {
  const fields: Record<string, string | number> = {};
  for (const col of columns) {
    if (!col.isPhoto) {
      fields[col.name] = col.name.toLowerCase() === 'status' ? 'offen' : '';
    }
  }
  return fields;
}

export function protocolStats(protocol: SiteReportProtocol) {
  const photoCount = protocol.entries.filter((e) => e.photoPath).length;
  const openCount = protocol.entries.filter((e) => {
    const statusCol = protocol.columns.find((c) => c.name.toLowerCase() === 'status');
    if (!statusCol) return false;
    return String(e.fields[statusCol.name] ?? '').toLowerCase() === 'offen';
  }).length;
  return {
    entryCount: protocol.entries.length,
    photoCount,
    openCount
  };
}

export async function removeProtocolEntry(
  protocol: SiteReportProtocol,
  entryId: string
): Promise<SiteReportProtocol> {
  const entry = protocol.entries.find((row) => row.id === entryId);
  if (entry?.photoPath) {
    await deleteEntryPhoto(entry.photoPath);
  }
  const next: SiteReportProtocol = {
    ...protocol,
    entries: protocol.entries.filter((row) => row.id !== entryId)
  };
  await updateProtocol(next);
  return next;
}

export async function deleteProtocolWithCleanup(protocolId: string): Promise<void> {
  const exports = await listExports();
  for (const exp of exports.filter((row) => row.protocolId === protocolId)) {
    await deleteCachedExport(exp.id);
  }
  await deleteExportsByProtocol(protocolId);
  await deleteProtocolPhotos(protocolId);
  await softDeleteProtocol(protocolId);
}

export async function bulkDeleteProtocols(protocolIds: string[]): Promise<void> {
  for (const id of protocolIds) {
    await deleteProtocolWithCleanup(id);
  }
}

export async function upsertProtocolEntry(
  protocol: SiteReportProtocol,
  entry: SiteReportEntry,
  editingEntryId?: string | null
): Promise<SiteReportProtocol> {
  let entries: SiteReportEntry[];
  if (editingEntryId) {
    entries = protocol.entries.map((row) => (row.id === editingEntryId ? entry : row));
  } else {
    entries = [entry, ...protocol.entries];
  }
  const next = { ...protocol, entries };
  await updateProtocol(next);
  return next;
}

export async function getProtocolOrThrow(id: string): Promise<SiteReportProtocol> {
  const protocol = await getProtocol(id);
  if (!protocol) {
    throw new Error('Protokoll nicht gefunden.');
  }
  return protocol;
}
