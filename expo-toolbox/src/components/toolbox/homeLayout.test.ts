import assert from 'node:assert/strict';

import { resolveHomeLayoutTier } from './homeLayout';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('resolveHomeLayoutTier picks dense on short screens', () => {
  assert.equal(resolveHomeLayoutTier(520), 'dense');
});

test('resolveHomeLayoutTier picks compact on medium screens', () => {
  assert.equal(resolveHomeLayoutTier(620), 'compact');
});

test('resolveHomeLayoutTier picks relaxed on tall screens', () => {
  assert.equal(resolveHomeLayoutTier(820), 'relaxed');
});

console.log(`\n${passed} tests passed`);
