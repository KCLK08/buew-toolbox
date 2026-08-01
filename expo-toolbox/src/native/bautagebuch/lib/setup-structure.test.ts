import assert from 'node:assert/strict';

import {
  addStructureGroup,
  addStructureTable,
  completeStructureStep,
  deleteStructureItem,
  getStructureItems,
  migrateStructureFromLegacy,
  moveStructureItem,
  normalizeStructureTableColumns,
  updateStructureGroup,
  updateStructureTable,
  validateStructureStep
} from './setup-structure';
import { getWizardState, withWizardState } from './setup-mapping';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('getStructureItems returns empty list for fresh wizard', () => {
  const model = withWizardState({ single_sections: [] }, { step: 'structure', structure: [] });
  assert.deepEqual(getStructureItems(model), []);
});

test('addStructureGroup appends group with order', () => {
  let model: Record<string, unknown> = { wizard: { step: 'structure', structure: [], groups: [], tables: [] } };
  model = addStructureGroup(model, { name: 'Allgemeine Angaben', description: 'Projektinfos' });
  const items = getStructureItems(model);
  assert.equal(items.length, 1);
  assert.equal(items[0].type, 'group');
  if (items[0].type === 'group') {
    assert.equal(items[0].name, 'Allgemeine Angaben');
    assert.equal(items[0].description, 'Projektinfos');
    assert.equal(items[0].order, 0);
  }
  const wizard = getWizardState(model);
  assert.equal(wizard.groups.length, 1);
  assert.equal(wizard.groups[0].label, 'Allgemeine Angaben');
});

test('addStructureTable stores logical columns', () => {
  let model: Record<string, unknown> = { wizard: { step: 'structure', structure: [], groups: [], tables: [] } };
  model = addStructureTable(model, {
    name: 'Arbeitsleistungen',
    columns: [{ name: 'Tätigkeit' }, { name: 'Menge' }]
  });
  const items = getStructureItems(model);
  assert.equal(items.length, 1);
  assert.equal(items[0].type, 'table');
  if (items[0].type === 'table') {
    assert.equal(items[0].columns.length, 2);
    assert.equal(items[0].columns[0].name, 'Tätigkeit');
  }
});

test('addStructureTable preserves planned column slots with default names', () => {
  let model: Record<string, unknown> = { wizard: { step: 'structure', structure: [], groups: [], tables: [] } };
  model = addStructureTable(model, {
    name: 'Arbeitskräfte',
    columns: [{ name: 'Mitarbeiter' }, { name: '' }, { name: '' }, { name: 'Ende' }]
  });
  const table = getStructureItems(model)[0];
  assert.equal(table.type, 'table');
  if (table.type !== 'table') return;
  assert.equal(table.columns.length, 4);
  assert.deepEqual(
    table.columns.map((column) => column.name),
    ['Mitarbeiter', 'Spalte 2', 'Spalte 3', 'Ende']
  );
});

test('normalizeStructureTableColumns fills unnamed slots', () => {
  assert.deepEqual(
    normalizeStructureTableColumns([{ name: 'Stunden' }, { name: '   ' }]).map((column) => column.name),
    ['Stunden', 'Spalte 2']
  );
});

test('addStructureTable allows name-only tables without columns', () => {
  let model: Record<string, unknown> = { wizard: { step: 'structure', structure: [], groups: [], tables: [] } };
  model = addStructureTable(model, { name: 'Arbeitskräfte', columns: [] });
  const table = getStructureItems(model)[0];
  assert.equal(table.type, 'table');
  if (table.type !== 'table') return;
  assert.equal(table.name, 'Arbeitskräfte');
  assert.equal(table.columns.length, 0);
});

test('moveStructureItem reorders entries', () => {
  let model: Record<string, unknown> = { wizard: { step: 'structure', structure: [], groups: [], tables: [] } };
  model = addStructureGroup(model, { name: 'A' });
  model = addStructureTable(model, { name: 'B', columns: [{ name: 'Spalte' }] });
  const firstId = getStructureItems(model)[0].id;
  model = moveStructureItem(model, firstId, 1);
  const items = getStructureItems(model);
  assert.equal(items[0].name, 'B');
  assert.equal(items[1].name, 'A');
});

test('validateStructureStep requires at least one item', () => {
  assert.match(String(validateStructureStep([])), /mindestens/);
});

test('completeStructureStep advances wizard to assign and builds shells', () => {
  let model: Record<string, unknown> = { wizard: { step: 'structure', structure: [], groups: [], tables: [] } };
  model = addStructureGroup(model, { name: 'Wetter' });
  model = completeStructureStep(model);
  const wizard = getWizardState(model);
  assert.equal(wizard.step, 'assign');
  assert.equal(Array.isArray(model.single_sections) ? model.single_sections.length : 0, 1);
  assert.equal(Array.isArray(model.table_sections) ? model.table_sections.length : 0, 0);
});

test('completeStructureStep preserves assign progress when returning from step 2', () => {
  let model: Record<string, unknown> = { wizard: { step: 'structure', structure: [], groups: [], tables: [] } };
  model = addStructureGroup(model, { name: 'Wetter' });
  const groupId = getStructureItems(model)[0].id;
  model = completeStructureStep(model);
  model = withWizardState(model, {
    step: 'assign',
    assignments: { f1: groupId, f2: groupId },
    currentFieldIndex: 7,
    fieldLabels: { f1: 'Datum' }
  });
  model = withWizardState(model, { step: 'structure' });
  model = completeStructureStep(model);
  const wizard = getWizardState(model);
  assert.equal(wizard.step, 'assign');
  assert.deepEqual(wizard.assignments, { f1: groupId, f2: groupId });
  assert.equal(wizard.currentFieldIndex, 7);
  assert.deepEqual(wizard.fieldLabels, { f1: 'Datum' });
});

test('mixed group/table order is preserved in section_order', () => {
  let model: Record<string, unknown> = { wizard: { step: 'structure', structure: [], groups: [], tables: [] } };
  model = addStructureGroup(model, { name: 'Allgemein' });
  model = addStructureTable(model, { name: 'Arbeit', columns: [{ name: 'Tätigkeit' }] });
  model = addStructureGroup(model, { name: 'Wetter' });
  model = completeStructureStep(model);
  const order = model.section_order as Array<{ kind: string }>;
  assert.deepEqual(
    order.map((entry) => entry.kind),
    ['single', 'table', 'single']
  );
});

test('updateStructureGroup and updateStructureTable persist edits', () => {
  let model: Record<string, unknown> = { wizard: { step: 'structure', structure: [], groups: [], tables: [] } };
  model = addStructureGroup(model, { name: 'Alt' });
  const groupId = getStructureItems(model)[0].id;
  model = updateStructureGroup(model, groupId, { name: 'Neu', description: 'Info' });
  const group = getStructureItems(model)[0];
  assert.equal(group.type, 'group');
  if (group.type === 'group') {
    assert.equal(group.name, 'Neu');
    assert.equal(group.description, 'Info');
  }

  model = addStructureTable(model, { name: 'Tabelle', columns: [{ name: 'A' }] });
  const table = getStructureItems(model).find((item) => item.type === 'table');
  assert.ok(table && table.type === 'table');
  if (!table || table.type !== 'table') return;
  model = updateStructureTable(model, table.id, {
    name: 'Tabelle 2',
    columns: [{ id: table.columns[0].id, name: 'B' }]
  });
  const updated = getStructureItems(model).find((item) => item.id === table.id);
  assert.ok(updated && updated.type === 'table');
  if (!updated || updated.type !== 'table') return;
  assert.equal(updated.name, 'Tabelle 2');
  assert.equal(updated.columns[0].name, 'B');
});

test('deleteStructureItem removes entry from wizard mirrors', () => {
  let model: Record<string, unknown> = { wizard: { step: 'structure', structure: [], groups: [], tables: [] } };
  model = addStructureGroup(model, { name: 'Temp' });
  const id = getStructureItems(model)[0].id;
  model = deleteStructureItem(model, id);
  assert.equal(getStructureItems(model).length, 0);
  assert.equal(getWizardState(model).groups.length, 0);
});

test('migrateStructureFromLegacy builds ordered items from groups and tables', () => {
  const items = migrateStructureFromLegacy(
    [{ sectionId: 'g1', label: 'Gruppe' }],
    [
      {
        tableId: 't1',
        label: 'Tabelle',
        rowCount: 1,
        columns: [{ columnId: 'c1', label: 'Spalte', type: 'text' }]
      }
    ]
  );
  assert.equal(items.length, 2);
  assert.equal(items[0].name, 'Gruppe');
  assert.equal(items[1].name, 'Tabelle');
});

console.log(`\n${passed} setup-structure tests passed`);
