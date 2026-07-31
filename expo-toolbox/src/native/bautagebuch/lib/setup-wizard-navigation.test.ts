import assert from 'node:assert/strict';

import {
  prepareSetupStepNavigation,
  resolveSetupOverviewPath,
  resolveSetupStepPath
} from './setup-wizard-navigation';
import { getWizardState, withWizardState } from './setup-mapping';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('resolveSetupOverviewPath points to template setup tab', () => {
  assert.equal(String(resolveSetupOverviewPath()), '/bautagebuch/config');
});

test('prepareSetupStepNavigation updates wizard step without resetting state', () => {
  const source = withWizardState(
    { single_sections: [{ sectionId: 'g1', fields: [] }] },
    { step: 'assign', currentFieldIndex: 2, assignments: { f1: 'g1' } }
  );
  const next = prepareSetupStepNavigation(source, 'fields');
  const wizard = getWizardState(next);
  assert.equal(wizard.step, 'fields');
  assert.equal(wizard.currentFieldIndex, 2);
  assert.equal(wizard.assignments.f1, 'g1');
});

test('resolveSetupStepPath maps wizard steps to setup routes', () => {
  assert.equal(String(resolveSetupStepPath('tpl_1', 'structure')), '/bautagebuch/setup/tpl_1/mapping');
  assert.equal(String(resolveSetupStepPath('tpl_1', 'assign')), '/bautagebuch/setup/tpl_1/assign');
  assert.equal(String(resolveSetupStepPath('tpl_1', 'fields')), '/bautagebuch/setup/tpl_1/fields');
});

console.log(`\n${passed} setup-wizard-navigation tests passed`);
