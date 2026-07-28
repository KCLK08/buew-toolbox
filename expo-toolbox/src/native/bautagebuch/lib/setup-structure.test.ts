import assert from 'node:assert/strict';

import {
  addStructureGroup,
  addStructureTable,
  completeStructureStep,
  getStructureItems,
  moveStructureItem,
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

console.log(`\n${passed} setup-structure tests passed`);
