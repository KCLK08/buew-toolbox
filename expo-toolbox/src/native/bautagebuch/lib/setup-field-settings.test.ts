import assert from 'node:assert/strict';

import {
  applyFieldTypeChange,
  getFieldSettingsProgress,
  listFieldSettingsTargets,
  markFieldSettingsTargetConfigured,
  normalizeSetupFieldType,
  resolveCurrentFieldSettingsIndex
} from './setup-field-settings';
import { getWizardState, withWizardState } from './setup-mapping';
import type { SetupFieldConfig } from '../types';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

const setupModel = {
  single_sections: [
    {
      sectionId: 'g1',
      label: 'Wetter',
      fields: [
        {
          fieldId: 'f1',
          fieldName: 'Temp',
          label: 'Temperatur',
          type: 'text',
          page: 1
        } satisfies SetupFieldConfig
      ]
    }
  ],
  table_sections: [],
  section_order: [{ kind: 'single', id: 'g1' }],
  wizard: { step: 'fields', configuredFieldIds: [], currentFieldSettingsIndex: 0 }
};

test('listFieldSettingsTargets returns group fields in order', () => {
  const targets = listFieldSettingsTargets(setupModel);
  assert.equal(targets.length, 1);
  assert.equal(targets[0]?.kind, 'single');
  assert.equal(targets[0]?.fieldId, 'f1');
});

test('normalizeSetupFieldType maps dropdown detection to select', () => {
  const field: SetupFieldConfig = { fieldId: 'f1', type: 'dropdown' };
  assert.equal(normalizeSetupFieldType(field), 'select');
});

test('applyFieldTypeChange seeds select options from detection', () => {
  const field: SetupFieldConfig = { fieldId: 'f1', label: 'Wetter' };
  const patch = applyFieldTypeChange(field, 'select', [
    {
      id: 'd1',
      templateId: 't',
      fieldId: 'f1',
      fieldName: 'W1',
      labelCandidate: 'W1',
      type: 'dropdown',
      options: ['Sonne', 'Regen'],
      page: 1,
      orderIndex: 0,
      rect: null,
      geometry: null,
      source: 'acroform',
      createdAt: '',
      updatedAt: ''
    }
  ]);
  assert.deepEqual(patch.options, ['Sonne', 'Regen']);
});

test('field settings progress tracks configured targets', () => {
  const targets = listFieldSettingsTargets(setupModel);
  const wizard = getWizardState(setupModel);
  assert.equal(getFieldSettingsProgress(targets, wizard).open, 1);
  const marked = markFieldSettingsTargetConfigured(setupModel, targets[0]);
  assert.equal(getFieldSettingsProgress(targets, getWizardState(marked)).configured, 1);
});

test('resolveCurrentFieldSettingsIndex prefers first open target', () => {
  const targets = listFieldSettingsTargets(setupModel);
  const wizard = getWizardState(
    withWizardState(setupModel, {
      configuredFieldIds: [targets[0].key]
    })
  );
  assert.equal(resolveCurrentFieldSettingsIndex(targets, wizard), 0);
});

console.log(`\n${passed} tests passed`);
