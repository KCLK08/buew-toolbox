import assert from 'node:assert/strict';

import {
  previewScrollOverlayPlacement,
  resolvePreviewOverlayPlacement
} from './pdf-preview-overlay';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('resolvePreviewOverlayPlacement avoids center-band overlap via bottom-sheet', () => {
  assert.equal(resolvePreviewOverlayPlacement([200, 500, 280, 540]), 'bottom-sheet');
});

test('resolvePreviewOverlayPlacement maps top field to bottom panel slot', () => {
  assert.equal(resolvePreviewOverlayPlacement([100, 720, 200, 760]), 'bottom');
});

test('resolvePreviewOverlayPlacement maps bottom field to top panel slot', () => {
  assert.equal(resolvePreviewOverlayPlacement([100, 40, 200, 80]), 'top');
});

test('previewScrollOverlayPlacement maps bottom-sheet to bottom scroll bias', () => {
  assert.equal(previewScrollOverlayPlacement('bottom-sheet'), 'bottom');
});

console.log(`\n${passed} tests passed`);
