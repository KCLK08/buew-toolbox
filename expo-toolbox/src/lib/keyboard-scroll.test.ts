import assert from 'node:assert/strict';

import {
  resolveKeyboardScrollDelta,
  resolveKeyboardVisibleBounds
} from './keyboard-scroll';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('resolveKeyboardVisibleBounds uses keyboard screenY when open', () => {
  const bounds = resolveKeyboardVisibleBounds(800, { screenY: 520, height: 280 }, 48);
  assert.equal(bounds.bottom, 472);
  assert.equal(bounds.top, 0);
});

test('resolveKeyboardVisibleBounds uses full window when keyboard closed', () => {
  const bounds = resolveKeyboardVisibleBounds(800, null, 48);
  assert.equal(bounds.bottom, 752);
});

test('resolveKeyboardScrollDelta scrolls down when field is below keyboard', () => {
  const delta = resolveKeyboardScrollDelta(500, 44, 0, 472, 28);
  assert.equal(delta, 100);
});

test('resolveKeyboardScrollDelta scrolls up when field is above visible top', () => {
  const delta = resolveKeyboardScrollDelta(12, 44, 80, 472, 28);
  assert.equal(delta, -96);
});

test('resolveKeyboardScrollDelta returns zero when field is visible', () => {
  const delta = resolveKeyboardScrollDelta(200, 44, 0, 472, 28);
  assert.equal(delta, 0);
});

console.log(`\n${passed} keyboard-scroll tests passed`);
