import { syncSectionOrder } from './setup-model.js';

export type OrderedSectionEntry = {
  kind: 'single' | 'table';
  id: string;
  label: string;
  typeLabel: string;
};

export function listOrderedSections(setupModel: Record<string, unknown>): OrderedSectionEntry[] {
  const singleSections = (Array.isArray(setupModel.single_sections) ? setupModel.single_sections : []) as Array<{
    sectionId?: string;
    label?: string;
  }>;
  const tableSections = (Array.isArray(setupModel.table_sections) ? setupModel.table_sections : []) as Array<{
    tableId?: string;
    label?: string;
  }>;
  const sectionOrder = syncSectionOrder(setupModel);

  return sectionOrder.map((entry) => {
    if (entry.kind === 'single') {
      const section = singleSections.find((item) => String(item.sectionId) === entry.id);
      return {
        kind: 'single' as const,
        id: entry.id,
        label: String(section?.label || entry.id),
        typeLabel: 'Gruppe'
      };
    }
    const table = tableSections.find((item) => String(item.tableId) === entry.id);
    return {
      kind: 'table' as const,
      id: entry.id,
      label: String(table?.label || entry.id),
      typeLabel: 'Tabelle'
    };
  });
}

export function sectionEntryKey(entry: Pick<OrderedSectionEntry, 'kind' | 'id'>): string {
  return `${entry.kind}:${entry.id}`;
}

export function moveSectionInSetupModel(
  setupModel: Record<string, unknown>,
  index: number,
  direction: -1 | 1
): Record<string, unknown> {
  const order = syncSectionOrder(setupModel);
  const nextIndex = index + direction;
  if (nextIndex < 0 || nextIndex >= order.length) return setupModel;
  const next = [...order];
  [next[index], next[nextIndex]] = [next[nextIndex], next[index]];
  return {
    ...setupModel,
    section_order: next
  };
}
