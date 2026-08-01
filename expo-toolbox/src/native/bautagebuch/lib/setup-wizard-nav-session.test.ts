import assert from 'node:assert/strict';

import {
  clearSetupStepNavigation,
  consumeSetupStepNavigation,
  getSetupStepNavigationTarget,
  isSetupStepNavigationActive,
  markSetupStepNavigation
} from './setup-wizard-nav-session';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('markSetupStepNavigation records the target step', () => {
  clearSetupStepNavigation();
  markSetupStepNavigation('assign');
  assert.equal(getSetupStepNavigationTarget(), 'assign');
  assert.equal(isSetupStepNavigationActive(), true);
});

test('consumeSetupStepNavigation clears only when step matches', () => {
  clearSetupStepNavigation();
  markSetupStepNavigation('fields');
  assert.equal(consumeSetupStepNavigation('structure'), false);
  assert.equal(getSetupStepNavigationTarget(), 'fields');
  assert.equal(consumeSetupStepNavigation('fields'), true);
  assert.equal(getSetupStepNavigationTarget(), null);
  assert.equal(isSetupStepNavigationActive(), false);
});

test('clearSetupStepNavigation resets pending navigation', () => {
  markSetupStepNavigation('structure');
  clearSetupStepNavigation();
  assert.equal(getSetupStepNavigationTarget(), null);
  assert.equal(consumeSetupStepNavigation('structure'), false);
});

console.log(`\n${passed} setup-wizard-nav-session tests passed`);
