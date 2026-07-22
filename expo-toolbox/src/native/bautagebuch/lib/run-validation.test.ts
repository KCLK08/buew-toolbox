import assert from 'node:assert/strict';

import {
  computeTotalMissingRequired,
  isPhotoDocRequiredMissing,
  sectionRunOptions
} from './run-validation';

function headerSection() {
  return {
    sectionId: 'single:header',
    kind: 'single',
    label: 'Kopfdaten',
    fields: [
      { fieldId: 'f1', fieldName: 'Text1', label: 'Projekt', type: 'text', required: true },
      { fieldId: 'f2', fieldName: 'Text3', label: 'Gewerk A', type: 'text', required: false }
    ]
  };
}

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('computeTotalMissingRequired counts unfilled required fields', () => {
  const setupModel = { single_sections: [{ sectionId: 'header', fields: headerSection().fields }] };
  const missing = computeTotalMissingRequired(setupModel, {}, null);
  assert.equal(missing >= 2, true);
});

test('computeTotalMissingRequired passes when required fields and photo choice are set', () => {
  const setupModel = { single_sections: [{ sectionId: 'header', fields: headerSection().fields }] };
  const values = {
    'field:f1': 'Projekt Nord',
    'field:f2': 'X'
  };
  const missing = computeTotalMissingRequired(setupModel, values, false);
  assert.equal(missing, 0);
});

test('isPhotoDocRequiredMissing is true until Ja/Nein is chosen', () => {
  assert.equal(isPhotoDocRequiredMissing(null), true);
  assert.equal(isPhotoDocRequiredMissing(undefined), true);
  assert.equal(isPhotoDocRequiredMissing(false), false);
  assert.equal(isPhotoDocRequiredMissing(true), false);
});

test('sectionRunOptions exposes gewerk group for header sections', () => {
  const options = sectionRunOptions(headerSection() as never, {});
  assert.equal(Array.isArray(options.requiredAnyGroups), true);
  assert.equal(options.requiredAnyGroups?.some((group) => group.label === 'Gewerk'), true);
});

console.log(`\n${passed} tests passed`);
