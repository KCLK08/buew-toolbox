import assert from 'node:assert/strict';

import {
  assignFieldToGroup,
  assignFieldToTableCell,
  deferField,
  ensureWizardInitialized,
  getMappingProgress,
  getWizardState,
  isMappingComplete,
  rebuildSectionsFromWizard,
  resolveCurrentMappingIndex,
  resolveOverlayPlacement,
  resolveSetupEntryPath,
  checkMappingTransition,
  getMappingCompletionSummary,
  sortMappingFields,
  withWizardState
} from './setup-mapping';
import type { DetectedField } from '../types';

function field(id: string, page = 1): DetectedField {
  return {
    id: `df_${id}`,
    templateId: 'tpl_test',
    fieldId: id,
    fieldName: id,
    labelCandidate: id,
    type: 'text',
    options: [],
    page,
    orderIndex: 0,
    rect: [100, 200, 200, 240],
    createdAt: '',
    updatedAt: ''
  };
}

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('ensureWizardInitialized does not reset wizard step', () => {
  const model = withWizardState({ single_sections: [] }, { step: 'fields' });
  const next = ensureWizardInitialized(model);
  assert.equal(getWizardState(next).step, 'fields');
});

test('ensureWizardInitialized starts new templates without preset groups', () => {
  const next = ensureWizardInitialized({ single_sections: [] });
  assert.deepEqual(getWizardState(next).groups, []);
  assert.equal(getWizardState(next).step, 'structure');
});

test('getWizardState returns empty groups when none were saved', () => {
  const wizard = getWizardState({});
  assert.deepEqual(wizard.groups, []);
});

test('resolveSetupEntryPath routes wizard steps to setup screens', () => {
  assert.equal(
    String(resolveSetupEntryPath('tpl_1', withWizardState({}, { step: 'structure' }))),
    '/bautagebuch/setup/tpl_1/mapping'
  );
  assert.equal(
    String(resolveSetupEntryPath('tpl_1', withWizardState({}, { step: 'assign' }))),
    '/bautagebuch/setup/tpl_1/assign'
  );
  assert.equal(
    String(resolveSetupEntryPath('tpl_1', withWizardState({}, { step: 'fields' }))),
    '/bautagebuch/setup/tpl_1/fields'
  );
});

test('legacy mapping step with assignments resolves to assign', () => {
  const wizard = getWizardState({
    wizard: {
      step: 'assign',
      assignments: { a: 'g1' },
      groups: [{ sectionId: 'g1', label: 'G1' }]
    }
  });
  assert.equal(wizard.step, 'assign');
});

test('isMappingComplete treats deferred fields as handled', () => {
  const fields = sortMappingFields([field('a'), field('b')]);
  let model: Record<string, unknown> = { wizard: { groups: [{ sectionId: 'g1', label: 'G1' }] } };
  model = assignFieldToGroup(model, 'a', 'g1');
  model = deferField(model, 'b');
  assert.equal(isMappingComplete(fields, getWizardState(model)), true);
});

test('rebuildSectionsFromWizard keeps deferred fields in fallback group', () => {
  const fields = sortMappingFields([field('a'), field('b')]);
  let model = withWizardState(
    { single_sections: [] },
    {
      step: 'assign',
      groups: [
        { sectionId: 'kopfdaten', label: 'Kopfdaten' },
        { sectionId: 'sonstiges', label: 'Sonstiges' }
      ],
      assignments: { a: 'kopfdaten' },
      deferredFieldIds: ['b']
    }
  );
  const rebuilt = rebuildSectionsFromWizard(model, fields);
  const sections = rebuilt.single_sections as Array<{ sectionId: string; fields: Array<{ fieldId: string }> }>;
  const sonstiges = sections.find((section) => section.sectionId === 'sonstiges');
  assert.ok(sonstiges);
  assert.equal(sonstiges?.fields.some((entry) => entry.fieldId === 'b'), true);
});

test('getMappingProgress counts deferred fields as handled', () => {
  const fields = sortMappingFields([field('a'), field('b'), field('c')]);
  const wizard = getWizardState(
    withWizardState(
      {},
      {
        assignments: { a: 'kopfdaten' },
        deferredFieldIds: ['b']
      }
    )
  );
  const progress = getMappingProgress(fields, wizard);
  assert.equal(progress.percent, 67);
  assert.equal(progress.remaining, 1);
});

test('resolveOverlayPlacement keeps panel away from field edges', () => {
  assert.equal(resolveOverlayPlacement([100, 720, 200, 760]), 'bottom');
  assert.equal(resolveOverlayPlacement([100, 40, 200, 80]), 'top');
  assert.equal(resolveOverlayPlacement([20, 400, 80, 440]), 'right');
});

test('resolveCurrentMappingIndex prefers first unassigned field', () => {
  const fields = sortMappingFields([field('a'), field('b'), field('c')]);
  const model = withWizardState(
    {},
    {
      currentFieldIndex: 0,
      groups: [{ sectionId: 'g1', label: 'G1' }],
      assignments: { a: 'g1' }
    }
  );
  const wizard = getWizardState(model);
  assert.equal(resolveCurrentMappingIndex(fields, wizard), 1);
});

test('checkMappingTransition reports deferred fields', () => {
  const fields = sortMappingFields([field('a'), field('b')]);
  let model = withWizardState(
    {},
    {
      groups: [{ sectionId: 'g1', label: 'G1' }],
      assignments: { a: 'g1' },
      deferredFieldIds: ['b']
    }
  );
  const check = checkMappingTransition(model, fields);
  assert.equal(check.unassignedCount, 1);
  assert.ok(check.issues.some((issue) => issue.includes('1 Felder')));
});

test('getMappingCompletionSummary aggregates group counts', () => {
  const fields = sortMappingFields([field('a'), field('b'), field('c')]);
  const model = withWizardState(
    {},
    {
      groups: [
        { sectionId: 'g1', label: 'Kopfdaten' },
        { sectionId: 'g2', label: 'Sonstiges' }
      ],
      assignments: { a: 'g1', b: 'g1', c: 'g2' }
    }
  );
  const summary = getMappingCompletionSummary(model, fields);
  assert.equal(summary.totalFields, 3);
  assert.equal(summary.groupCount, 2);
  assert.equal(summary.groups.find((g) => g.sectionId === 'g1')?.fieldCount, 2);
});

test('isMappingComplete treats table assignments as handled', () => {
  const fields = sortMappingFields([field('a')]);
  const model = withWizardState(
    {},
    {
      groups: [{ sectionId: 'g1', label: 'G1' }],
      tables: [
        {
          tableId: 'tbl_1',
          label: 'Personal',
          rowCount: 1,
          columns: [{ columnId: 'c1', label: 'Checkbox', type: 'checkbox' }]
        }
      ],
      tableAssignments: {
        a: { tableId: 'tbl_1', rowIndex: 0, columnId: 'c1' }
      }
    }
  );
  assert.equal(isMappingComplete(fields, getWizardState(model)), true);
});

test('rebuildSectionsFromWizard preserves field options and builds table sections', () => {
  const checkboxField: DetectedField = {
    ...field('cb1'),
    type: 'checkbox',
    options: ['Ja', 'Nein']
  };
  const fields = sortMappingFields([checkboxField]);
  let model = withWizardState(
    { single_sections: [] },
    {
      step: 'assign',
      groups: [{ sectionId: 'sonstiges', label: 'Sonstiges' }],
      tables: [
        {
          tableId: 'tbl_1',
          label: 'Checkliste',
          rowCount: 1,
          columns: [{ columnId: 'c1', label: 'Erledigt', type: 'checkbox' }]
        }
      ],
      tableAssignments: {
        cb1: { tableId: 'tbl_1', rowIndex: 0, columnId: 'c1' }
      }
    }
  );
  model = assignFieldToTableCell(model, 'cb1', { tableId: 'tbl_1', rowIndex: 0, columnId: 'c1' });
  const rebuilt = rebuildSectionsFromWizard(model, fields);
  const tables = rebuilt.table_sections as Array<{
    tableId: string;
    rows: Array<{ cells: Array<{ fieldId: string; options?: string[]; type?: string }> }>;
  }>;
  assert.equal(tables.length, 1);
  assert.equal(tables[0]?.rows[0]?.cells[0]?.fieldId, 'cb1');
  assert.equal(tables[0]?.rows[0]?.cells[0]?.type, 'checkbox');
  assert.deepEqual(tables[0]?.rows[0]?.cells[0]?.options, ['Ja', 'Nein']);
});

console.log(`\n${passed} tests passed`);
