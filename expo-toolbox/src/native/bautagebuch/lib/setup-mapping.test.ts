import assert from 'node:assert/strict';

import {
  assignFieldToGroup,
  assignFieldToTableCell,
  assignFieldToTableColumn,
  deferField,
  ensureWizardInitialized,
  getMappingProgress,
  getWizardState,
  isMappingComplete,
  markStructureIntroSeen,
  rebuildSectionsFromWizard,
  removeFieldFromWizard,
  resolveCurrentMappingIndex,
  resolveWalkthroughMappingIndex,
  resolveFieldAssignmentSummary,
  resolveOverlayPlacement,
  resolveSetupEntryPath,
  shouldShowStructureIntro,
  markAssignIntroSeen,
  markFieldsIntroSeen,
  shouldShowAssignIntro,
  shouldShowFieldsIntro,
  checkMappingTransition,
  getMappingCompletionSummary,
  resolveFieldDisplayLabel,
  resolveFieldEditLabel,
  sortMappingFields,
  buildFieldPreviewHighlights,
  withWizardState
} from './setup-mapping';
import { listFieldSettingsTargets } from './setup-field-settings';
import type {
  DetectedField,
  SetupFieldConfig
} from '../types';

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
    geometry: {
      page,
      rect: { x: 100, y: 200, width: 100, height: 40 }
    },
    source: 'acroform',
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
    '/bautagebuch/setup/tpl_1/intro'
  );
  assert.equal(
    String(
      resolveSetupEntryPath(
        'tpl_1',
        markStructureIntroSeen(withWizardState({}, { step: 'structure' }))
      )
    ),
    '/bautagebuch/setup/tpl_1/mapping'
  );
  assert.equal(
    String(resolveSetupEntryPath('tpl_1', withWizardState({}, { step: 'assign' }))),
    '/bautagebuch/setup/tpl_1/assign-intro'
  );
  assert.equal(
    String(
      resolveSetupEntryPath(
        'tpl_1',
        markAssignIntroSeen(withWizardState({}, { step: 'assign' }))
      )
    ),
    '/bautagebuch/setup/tpl_1/assign'
  );
  assert.equal(
    String(resolveSetupEntryPath('tpl_1', withWizardState({}, { step: 'fields' }))),
    '/bautagebuch/setup/tpl_1/fields-intro'
  );
  assert.equal(
    String(
      resolveSetupEntryPath(
        'tpl_1',
        markFieldsIntroSeen(withWizardState({}, { step: 'fields' }))
      )
    ),
    '/bautagebuch/setup/tpl_1/fields'
  );
});

test('shouldShowStructureIntro skips after intro seen or structure exists', () => {
  const fresh = withWizardState({}, { step: 'structure' });
  assert.equal(shouldShowStructureIntro(fresh), true);
  assert.equal(shouldShowStructureIntro(markStructureIntroSeen(fresh)), false);
  assert.equal(
    shouldShowStructureIntro(
      withWizardState(fresh, {
        structure: [{ id: 'g1', name: 'G', type: 'group', order: 0 }]
      })
    ),
    false
  );
});

test('shouldShowAssignIntro skips after intro seen or assignments exist', () => {
  const fresh = withWizardState({}, { step: 'assign' });
  assert.equal(shouldShowAssignIntro(fresh), true);
  assert.equal(shouldShowAssignIntro(markAssignIntroSeen(fresh)), false);
  assert.equal(
    shouldShowAssignIntro(
      withWizardState(fresh, {
        assignments: { f1: 'g1' }
      })
    ),
    false
  );
});

test('shouldShowFieldsIntro skips after intro seen or legacy templates', () => {
  const fresh = withWizardState({}, { step: 'fields' });
  assert.equal(shouldShowFieldsIntro(fresh), true);
  assert.equal(shouldShowFieldsIntro(markFieldsIntroSeen(fresh)), false);
  assert.equal(shouldShowFieldsIntro(fresh, 'builtin-etb'), false);
  assert.equal(
    shouldShowFieldsIntro(
      withWizardState(fresh, {
        structure: [{ id: 't1', name: 'T', type: 'table', order: 0, columns: [] }]
      }),
      ''
    ),
    true
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
  assert.equal(progress.percent, 33);
  assert.equal(progress.remaining, 1);
  assert.equal(progress.assigned, 1);
  assert.equal(progress.open, 2);
});

test('resolveOverlayPlacement keeps panel away from field edges', () => {
  assert.equal(resolveOverlayPlacement([100, 720, 200, 760]), 'bottom');
  assert.equal(resolveOverlayPlacement([100, 40, 200, 80]), 'top');
  assert.equal(resolveOverlayPlacement([20, 400, 80, 440]), 'right');
});

test('resolveCurrentMappingIndex respects stored index', () => {
  const fields = sortMappingFields([field('a'), field('b'), field('c')]);
  const model = withWizardState(
    {},
    {
      currentFieldIndex: 2,
      groups: [{ sectionId: 'g1', label: 'G1' }],
      assignments: { a: 'g1' }
    }
  );
  const wizard = getWizardState(model);
  assert.equal(resolveCurrentMappingIndex(fields, wizard), 2);
});

test('resolveWalkthroughMappingIndex prefers first unassigned field', () => {
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
  assert.equal(resolveWalkthroughMappingIndex(fields, wizard), 1);
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

test('isMappingComplete requires at least one field', () => {
  assert.equal(isMappingComplete([], getWizardState({ wizard: { groups: [] } })), false);
});

test('resolveFieldAssignmentSummary reports group and table assignments', () => {
  const model = withWizardState(
    {},
    {
      groups: [{ sectionId: 'g1', label: 'Allgemeine Angaben' }],
      tables: [
        {
          tableId: 't1',
          label: 'Arbeitszeiten',
          rowCount: 1,
          columns: [{ columnId: 'c1', label: 'Stunden', type: 'text' }]
        }
      ],
      assignments: { f1: 'g1' },
      tableAssignments: { f2: { tableId: 't1', rowIndex: 0, columnId: 'c1' } }
    }
  );
  assert.deepEqual(resolveFieldAssignmentSummary(model, 'f1'), {
    kind: 'group',
    label: 'Allgemeine Angaben',
    sectionId: 'g1'
  });
  assert.equal(resolveFieldAssignmentSummary(model, 'f2').kind, 'table');
  assert.equal(resolveFieldAssignmentSummary(model, 'missing').kind, 'none');
});

test('removeFieldFromWizard clears assignments and labels', () => {
  const model = withWizardState(
    {},
    {
      assignments: { f1: 'g1' },
      tableAssignments: { f2: { tableId: 't1', rowIndex: 0, columnId: 'c1' } },
      fieldLabels: { f1: 'Datum', f2: 'Stunden' },
      deferredFieldIds: ['f2'],
      currentFieldIndex: 2
    }
  );
  const next = removeFieldFromWizard(model, 'f1', 2);
  const wizard = getWizardState(next);
  assert.equal(wizard.assignments.f1, undefined);
  assert.equal(wizard.fieldLabels?.f1, undefined);
  assert.equal(wizard.currentFieldIndex, 1);
});

test('removeFieldFromWizard clears single_sections and table_sections field references', () => {
  const model = withWizardState(
    {
      single_sections: [
        {
          sectionId: 'g1',
          label: 'Kopf',
          fields: [
            { fieldId: 'f1', label: 'Datum', type: 'datetime' },
            { fieldId: 'f2', label: 'Baustelle', type: 'text' }
          ]
        }
      ],
      table_sections: [
        {
          tableId: 't1',
          label: 'Leistungen',
          columns: [{ columnId: 'c1', label: 'Checkbox', type: 'checkbox' }],
          rows: [
            {
              rowId: 'row_1',
              index: 1,
              cells: [
                {
                  cellId: 't1_row_1_c1',
                  tableId: 't1',
                  rowId: 'row_1',
                  columnId: 'c1',
                  fieldId: 'f3',
                  fieldName: 'f3',
                  label: 'Erledigt',
                  type: 'checkbox',
                  skipped: false
                }
              ]
            }
          ]
        }
      ]
    },
    {
      assignments: { f1: 'g1', f2: 'g1' },
      tableAssignments: { f3: { tableId: 't1', rowIndex: 0, columnId: 'c1' } },
      fieldLabels: { f1: 'Datum', f2: 'Baustelle', f3: 'Erledigt' },
      configuredFieldIds: ['single:g1:f1', 'table-cell:t1:f3']
    }
  );

  const afterSingle = removeFieldFromWizard(model, 'f1', 3);
  const singleFields = (
    (afterSingle.single_sections as Array<{ fields: Array<{ fieldId: string }> }>)[0]?.fields || []
  ).map((field) => field.fieldId);
  assert.deepEqual(singleFields, ['f2']);
  assert.equal(getWizardState(afterSingle).configuredFieldIds?.includes('single:g1:f1'), false);

  const afterTable = removeFieldFromWizard(model, 'f3', 3);
  const cell = (
    (afterTable.table_sections as Array<{ rows: Array<{ cells: Array<{ fieldId: string; skipped?: boolean }> }> }>)[0]
      ?.rows[0]?.cells[0]
  );
  assert.equal(cell?.fieldId, '');
  assert.equal(cell?.skipped, true);
  assert.equal(getWizardState(afterTable).configuredFieldIds?.includes('table-cell:t1:f3'), false);

  const targets = listFieldSettingsTargets(afterTable);
  assert.equal(
    targets.some(
      (target) => target.kind === 'table-cell' && target.fieldId === 'f3'
    ),
    false
  );
  assert.equal(
    targets.some((target) => target.kind === 'single' && target.fieldId === 'f1'),
    true
  );
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

test('sortMappingFields assigns stable displayOrder independent of geometry', () => {
  const withGeometry = field('a');
  const withoutGeometry: DetectedField = {
    ...field('b'),
    geometry: null,
    rect: null
  };
  const sorted = sortMappingFields([withoutGeometry, withGeometry]);
  assert.equal(sorted.length, 2);
  assert.equal(sorted[0]?.fieldId, 'b');
  assert.equal(sorted[0]?.displayOrder, 1);
  assert.equal(sorted[1]?.fieldId, 'a');
  assert.equal(sorted[1]?.displayOrder, 2);
});

test('resolveFieldDisplayLabel prefers labelCandidate over wizard fieldLabels', () => {
  const fields = sortMappingFields([{ ...field('a'), labelCandidate: 'Datum' }]);
  const wizard = getWizardState({ wizard: { fieldLabels: { a: 'Alt' } } });
  assert.equal(resolveFieldDisplayLabel(fields[0]!, wizard), 'Datum');
});

test('resolveFieldEditLabel preserves empty draft while editing', () => {
  const fields = sortMappingFields([{ ...field('a'), labelCandidate: 'Datum' }]);
  const wizard = getWizardState({});
  assert.equal(resolveFieldEditLabel(fields[0]!, wizard, { a: '' }), '');
  assert.equal(resolveFieldDisplayLabel(fields[0]!, wizard, { a: '' }), 'Datum');
});

test('buildFieldPreviewHighlights uses displayOrder for overlay index', () => {
  const withoutGeometry: DetectedField = {
    ...field('b', 1),
    geometry: null,
    rect: null
  };
  const fields = sortMappingFields([withoutGeometry, field('a', 1)]);
  const highlights = buildFieldPreviewHighlights(fields, (entry) => entry.labelCandidate);
  assert.equal(highlights.length, 1);
  assert.equal(highlights[0]?.fieldId, 'a');
  assert.equal(highlights[0]?.index, 2);
});

test('assignFieldToGroup stores custom field label', () => {
  const detected = { ...field('a'), labelCandidate: 'Baumaßnahme' };
  const fields = sortMappingFields([detected]);
  const model = assignFieldToGroup(
    withWizardState({}, { groups: [{ sectionId: 'g1', label: 'Allgemein' }] }),
    'a',
    'g1',
    'Baumaßnahme'
  );
  const wizard = getWizardState(model);
  assert.equal(wizard.fieldLabels?.a, 'Baumaßnahme');
  const rebuilt = rebuildSectionsFromWizard(model, fields);
  const sections = rebuilt.single_sections as Array<{ fields: Array<{ label: string }> }>;
  assert.equal(sections[0]?.fields[0]?.label, 'Baumaßnahme');
});

test('assignFieldToTableColumn creates column and assigns field', () => {
  const detected = { ...field('t1'), labelCandidate: 'Tätigkeit' };
  const fields = sortMappingFields([detected]);
  let model = withWizardState(
    {},
    {
      groups: [],
      tables: [{ tableId: 'tbl_1', label: 'Arbeitsleistungen', rowCount: 1, columns: [] }],
      structure: [
        { id: 'tbl_1', name: 'Arbeitsleistungen', type: 'table', order: 0, columns: [] }
      ]
    }
  );
  model = assignFieldToTableColumn(model, 't1', 'tbl_1', {
    newColumnName: 'Tätigkeit',
    fieldLabel: 'Tätigkeit'
  });
  const wizard = getWizardState(model);
  assert.equal(wizard.tables[0]?.columns.length, 1);
  assert.equal(wizard.tableAssignments.t1?.columnId, wizard.tables[0]?.columns[0]?.columnId);
  const rebuilt = rebuildSectionsFromWizard(model, fields);
  const tables = rebuilt.table_sections as Array<{ rows: Array<{ cells: Array<{ label: string }> }> }>;
  assert.equal(tables[0]?.rows[0]?.cells[0]?.label, 'Tätigkeit');
});

console.log(`\n${passed} tests passed`);
