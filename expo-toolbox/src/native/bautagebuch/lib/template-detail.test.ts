import assert from 'node:assert/strict';

import {
  buildFieldDetailRows,
  buildTemplateDetailOverview,
  formatTemplateUpdatedAt,
  resolveFieldPositionLabel
} from './template-detail';
import {
  markSetupCompleted,
  resolveSetupEditStepPath,
  resolveTemplateOpenPath,
  shouldShowStructureIntro,
  withWizardState
} from './setup-mapping';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('resolveTemplateOpenPath routes ready templates to detail view', () => {
  const model = withWizardState({ single_sections: [] }, { step: 'fields', setupCompleted: true });
  assert.equal(
    String(resolveTemplateOpenPath('tpl_1', model, 'ready')),
    '/bautagebuch/setup/tpl_1/detail'
  );
});

test('resolveTemplateOpenPath keeps draft templates in setup wizard', () => {
  const model = withWizardState({ single_sections: [] }, { step: 'structure', structure: [] });
  assert.equal(
    String(resolveTemplateOpenPath('tpl_1', model, 'draft')),
    '/bautagebuch/setup/tpl_1/intro'
  );
});

test('shouldShowStructureIntro is skipped after setup completion', () => {
  const model = withWizardState({ single_sections: [] }, { step: 'structure', setupCompleted: true });
  assert.equal(shouldShowStructureIntro(model), false);
});

test('resolveSetupEditStepPath opens mapping for structure edits', () => {
  assert.equal(
    String(resolveSetupEditStepPath('tpl_1', 'structure')),
    '/bautagebuch/setup/tpl_1/mapping'
  );
});

test('markSetupCompleted flags wizard as finished', () => {
  const model = markSetupCompleted(withWizardState({}, { editMode: true, step: 'fields' }));
  const wizard = model.wizard as { setupCompleted?: boolean; editMode?: boolean };
  assert.equal(wizard.setupCompleted, true);
  assert.equal(wizard.editMode, false);
});

test('buildTemplateDetailOverview lists groups and tables in structure order', () => {
  const setupModel = {
    updatedAt: '2026-07-29T12:00:00.000Z',
    section_order: [
      { kind: 'single', id: 'g1' },
      { kind: 'table', id: 't1' }
    ],
    single_sections: [
      {
        sectionId: 'g1',
        label: 'Allgemeine Angaben',
        fields: [{ fieldId: 'f1', label: 'Baustelle', type: 'text', required: true }]
      }
    ],
    table_sections: [
      {
        tableId: 't1',
        label: 'Arbeitsleistungen',
        columns: [{ columnId: 'c1', label: 'Tätigkeit', type: 'text' }],
        rows: []
      }
    ],
    wizard: {
      step: 'fields',
      setupCompleted: true,
      structure: [
        { id: 'g1', name: 'Allgemeine Angaben', type: 'group', order: 0 },
        {
          id: 't1',
          name: 'Arbeitsleistungen',
          type: 'table',
          order: 1,
          columns: [{ id: 'c1', name: 'Tätigkeit', order: 0 }]
        }
      ]
    }
  };

  const overview = buildTemplateDetailOverview(setupModel);
  assert.equal(overview.items.length, 2);
  assert.equal(overview.items[0].kind, 'group');
  if (overview.items[0].kind === 'group') {
    assert.equal(overview.items[0].fieldCount, 1);
    assert.equal(overview.items[0].fields[0].label, 'Baustelle');
  }
  assert.equal(overview.items[1].kind, 'table');
});

test('resolveFieldPositionLabel maps rects to readable positions', () => {
  assert.equal(resolveFieldPositionLabel([400, 700, 500, 780]), 'oben rechts');
});

test('formatTemplateUpdatedAt renders german date', () => {
  assert.equal(formatTemplateUpdatedAt('2026-07-29T12:00:00.000Z'), '29.07.2026');
});

test('buildFieldDetailRows includes datetime settings', () => {
  const rows = buildFieldDetailRows({
    fieldId: 'f1',
    label: 'Datum',
    type: 'datetime',
    required: true,
    useCurrentDate: true,
    dateMode: 'date',
    page: 1,
    rect: [400, 700, 500, 780]
  });
  assert.ok(rows.some((row) => row.label === 'Feldtyp' && row.value.includes('Datum')));
  assert.ok(rows.some((row) => row.label === 'Heutiges Datum übernehmen' && row.checked === true));
});

console.log(`\n${passed} template-detail tests passed`);
