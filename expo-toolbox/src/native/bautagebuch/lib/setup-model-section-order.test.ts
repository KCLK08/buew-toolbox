import assert from 'node:assert/strict';

import { buildLegacySectionOrder, buildRunSections, syncSectionOrder } from './setup-model.js';

function sampleModel() {
  return {
    single_sections: [
      {
        sectionId: 'header',
        label: 'Kopfdaten',
        fields: [{ fieldId: 'f1', label: 'Projekt', skipped: false }]
      },
      {
        sectionId: 'weather',
        label: 'Witterung',
        fields: [{ fieldId: 'f2', label: 'Wetter', skipped: false }]
      }
    ],
    table_sections: [
      {
        tableId: 'table_main_personal',
        label: 'Besetzung',
        columns: [{ columnId: 'c1', label: 'Name', skipped: false }],
        rows: [
          {
            rowId: 'r1',
            index: 1,
            cells: [{ columnId: 'c1', cellId: 'cell1', fieldId: 'f3', skipped: false }]
          }
        ]
      },
      {
        tableId: 'table_detail_blocks',
        label: 'Leistung',
        columns: [{ columnId: 'c1', label: 'Block', skipped: false }],
        rows: [
          {
            rowId: 'r1',
            index: 1,
            cells: [{ columnId: 'c1', cellId: 'cell2', fieldId: 'f4', skipped: false }]
          }
        ]
      }
    ]
  };
}

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('buildLegacySectionOrder lists singles before tables', () => {
  const order = buildLegacySectionOrder(sampleModel());
  assert.deepEqual(
    order.map((entry) => `${entry.kind}:${entry.id}`),
    ['single:header', 'single:weather', 'table:table_main_personal', 'table:table_detail_blocks']
  );
});

test('buildRunSections respects custom section_order', () => {
  const model = {
    ...sampleModel(),
    section_order: [
      { kind: 'single', id: 'header' },
      { kind: 'table', id: 'table_main_personal' },
      { kind: 'single', id: 'weather' },
      { kind: 'table', id: 'table_detail_blocks' }
    ]
  };
  const sections = buildRunSections(model);
  assert.deepEqual(
    sections.map((section) => section.sectionId),
    [
      'single:header',
      'table:table_main_personal',
      'single:weather',
      'table:table_detail_blocks'
    ]
  );
});

test('syncSectionOrder appends newly added sections', () => {
  const synced = syncSectionOrder({
    section_order: [{ kind: 'single', id: 'header' }],
    single_sections: sampleModel().single_sections,
    table_sections: sampleModel().table_sections
  });
  assert.equal(synced.length, 4);
  assert.equal(synced[0].id, 'header');
  assert.equal(synced[synced.length - 1].kind, 'table');
});

console.log(`\n${passed} tests passed`);
