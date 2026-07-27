import assert from 'node:assert/strict';

import { listSetupSections } from './setup-mapping';
import { listOrderedSections, moveSectionInSetupModel } from './setup-section-order';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('listSetupSections follows section_order', () => {
  const setupModel = {
    single_sections: [
      { sectionId: 'a', label: 'Alpha', fields: [{ fieldId: 'f1' }] },
      { sectionId: 'b', label: 'Beta', fields: [{ fieldId: 'f2' }] }
    ],
    section_order: [
      { kind: 'single', id: 'b' },
      { kind: 'single', id: 'a' }
    ]
  };

  const sections = listSetupSections(setupModel);
  assert.equal(sections[0]?.sectionId, 'b');
  assert.equal(sections[1]?.sectionId, 'a');
});

test('moveSectionInSetupModel swaps section_order entries', () => {
  const setupModel = {
    single_sections: [
      { sectionId: 'a', label: 'Alpha', fields: [] },
      { sectionId: 'b', label: 'Beta', fields: [] }
    ],
    section_order: [
      { kind: 'single', id: 'a' },
      { kind: 'single', id: 'b' }
    ]
  };

  const next = moveSectionInSetupModel(setupModel, 0, 1);
  const ordered = listOrderedSections(next);
  assert.equal(ordered[0]?.id, 'b');
  assert.equal(ordered[1]?.id, 'a');
});

console.log(`\n${passed} tests passed`);
