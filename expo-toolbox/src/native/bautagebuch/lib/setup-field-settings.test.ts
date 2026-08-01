import assert from 'node:assert/strict';

import {
  applyFieldTypeChange,
  advanceFieldSettingsWalkthrough,
  buildFieldLabelResolver,
  getFieldSettingsProgress,
  listFieldSettingsTargets,
  markFieldSettingsTargetConfigured,
  normalizeSetupFieldType,
  resolveCurrentFieldSettingsIndex,
  resolveWalkthroughFieldSettingsIndex,
  resolveFieldDisplayOrder,
  resolveHybridFieldLabel,
  resolveHybridFieldSource
} from './setup-field-settings';
import { getWizardState, sortMappingFields, withWizardState } from './setup-mapping';
import type { DetectedField, SetupFieldConfig } from '../types';

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

test('resolveCurrentFieldSettingsIndex respects stored index', () => {
  const targets = listFieldSettingsTargets(setupModel);
  const wizard = getWizardState(
    withWizardState(setupModel, {
      currentFieldSettingsIndex: 0,
      configuredFieldIds: [targets[0].key]
    })
  );
  assert.equal(resolveCurrentFieldSettingsIndex(targets, wizard), 0);
});

test('advanceFieldSettingsWalkthrough advances sequentially through all targets', () => {
  const targets = listFieldSettingsTargets(setupModel);
  const next = advanceFieldSettingsWalkthrough(setupModel, targets, targets[0]);
  const wizard = getWizardState(next);
  assert.equal(wizard.currentFieldSettingsIndex, 0);
  assert.equal(wizard.configuredFieldIds?.length, 1);
});

test('resolveWalkthroughFieldSettingsIndex keeps stored index for configured target', () => {
  const targets = listFieldSettingsTargets(setupModel);
  const wizard = getWizardState(
    withWizardState(setupModel, {
      currentFieldSettingsIndex: 0,
      configuredFieldIds: [targets[0].key]
    })
  );
  assert.equal(resolveWalkthroughFieldSettingsIndex(targets, wizard), 0);
});

test('resolveHybridFieldLabel prefers section label over mapping labelCandidate', () => {
  const targets = listFieldSettingsTargets(setupModel);
  const field = targets[0];
  const config = setupModel.single_sections[0].fields[0] as SetupFieldConfig;
  const mappingFields = sortMappingFields([
    {
      id: 'd1',
      templateId: 't',
      fieldId: 'f1',
      fieldName: 'Temp',
      labelCandidate: 'Alt',
      type: 'text',
      options: [],
      page: 1,
      orderIndex: 0,
      rect: null,
      geometry: null,
      source: 'acroform',
      createdAt: '',
      updatedAt: ''
    } satisfies DetectedField
  ]);
  const label = resolveHybridFieldLabel(setupModel, targets[0], config, mappingFields);
  assert.equal(label, 'Temperatur');
});

test('resolveHybridFieldSource falls back to detected field source', () => {
  const config: SetupFieldConfig = { fieldId: 'f1', label: 'Temperatur', type: 'text' };
  const source = resolveHybridFieldSource(config, [
    {
      id: 'd1',
      templateId: 't',
      fieldId: 'f1',
      fieldName: 'Temp',
      labelCandidate: 'Temperatur',
      type: 'text',
      options: [],
      page: 1,
      orderIndex: 0,
      rect: null,
      geometry: null,
      source: 'manual',
      createdAt: '',
      updatedAt: ''
    }
  ]);
  assert.equal(source, 'manual');
});

test('buildFieldLabelResolver uses section labels for PDF overlay', () => {
  const targets = listFieldSettingsTargets(setupModel);
  const mappingFields = sortMappingFields([
    {
      id: 'd1',
      templateId: 't',
      fieldId: 'f1',
      fieldName: 'Temp',
      labelCandidate: 'DB Name',
      type: 'text',
      options: [],
      page: 1,
      orderIndex: 0,
      rect: null,
      geometry: null,
      source: 'acroform',
      createdAt: '',
      updatedAt: ''
    } satisfies DetectedField
  ]);
  const resolveLabel = buildFieldLabelResolver(setupModel, targets, mappingFields);
  assert.equal(resolveLabel(mappingFields[0]), 'Temperatur');
});

test('resolveFieldDisplayOrder returns mapping displayOrder', () => {
  const mappingFields = sortMappingFields([
    {
      id: 'd1',
      templateId: 't',
      fieldId: 'f1',
      fieldName: 'A',
      labelCandidate: 'A',
      type: 'text',
      options: [],
      page: 1,
      orderIndex: 0,
      rect: null,
      geometry: null,
      source: 'acroform',
      createdAt: '',
      updatedAt: ''
    },
    {
      id: 'd2',
      templateId: 't',
      fieldId: 'f2',
      fieldName: 'B',
      labelCandidate: 'B',
      type: 'text',
      options: [],
      page: 1,
      orderIndex: 1,
      rect: null,
      geometry: null,
      source: 'manual',
      createdAt: '',
      updatedAt: ''
    } satisfies DetectedField
  ]);
  assert.equal(resolveFieldDisplayOrder(mappingFields, 'f2'), 2);
});

console.log(`\n${passed} tests passed`);
