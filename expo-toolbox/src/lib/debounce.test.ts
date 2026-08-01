import assert from 'node:assert/strict';

import { debounce } from './debounce';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('debounce flush runs pending callback immediately', () => {
  let value = '';
  const persist = debounce((next: string) => {
    value = next;
  }, 200);

  persist('first');
  assert.equal(value, '');
  persist.flush();
  assert.equal(value, 'first');
});

test('debounce flush is noop when nothing is pending', () => {
  let calls = 0;
  const persist = debounce(() => {
    calls += 1;
  }, 200);

  persist.flush();
  assert.equal(calls, 0);
});

test('debounce cancel drops pending callback', () => {
  let value = '';
  const persist = debounce((next: string) => {
    value = next;
  }, 200);

  persist('lost');
  persist.cancel();
  persist.flush();
  assert.equal(value, '');
});

console.log(`\n${passed} debounce tests passed`);
