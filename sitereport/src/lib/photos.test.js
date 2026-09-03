import { test } from 'node:test';
import assert from 'node:assert/strict';
import { layoutPhotoCollage, normalizeEntryPhotos, preferredPhotoColumns } from './photos.js';

test('normalizeEntryPhotos reads legacy photoBlob and photoBlobs', () => {
  assert.deepEqual(normalizeEntryPhotos({ photoBlob: 'a' }), ['a']);
  assert.deepEqual(normalizeEntryPhotos({ photoBlobs: ['a', 'b'] }), ['a', 'b']);
  assert.deepEqual(normalizeEntryPhotos({ photoBlobs: ['a'], photoBlob: 'legacy' }), ['a']);
  assert.deepEqual(normalizeEntryPhotos({}), []);
});

test('preferredPhotoColumns keeps few images in one row', () => {
  assert.equal(preferredPhotoColumns(1), 1);
  assert.equal(preferredPhotoColumns(2), 2);
  assert.equal(preferredPhotoColumns(3), 3);
  assert.equal(preferredPhotoColumns(4), 2);
});

test('layoutPhotoCollage fits all images into the given page box', () => {
  const sizes = [
    { width: 1600, height: 900 },
    { width: 900, height: 1600 },
    { width: 1200, height: 1200 }
  ];
  const layout = layoutPhotoCollage(sizes, 500, 260, { gap: 4, frame: 2 });
  assert.equal(layout.items.length, 3);
  assert.ok(layout.width <= 500 + 0.01);
  assert.ok(layout.height <= 260 + 0.01);
  assert.equal(layout.cols, 3);
  assert.equal(layout.rows, 1);
  for (const item of layout.items) {
    assert.ok(item.frameX >= 0);
    assert.ok(item.frameY >= 0);
    assert.ok(item.frameX + item.frameW <= layout.width + 0.01);
    assert.ok(item.frameY + item.frameH <= layout.height + 0.01);
  }
});

test('layoutPhotoCollage keeps up to three images in one row', () => {
  const sizes = [
    { width: 400, height: 300 },
    { width: 400, height: 300 },
    { width: 400, height: 300 }
  ];
  const layout = layoutPhotoCollage(sizes, 500, 500, { gap: 4, frame: 2 });
  assert.equal(layout.cols, 3);
  assert.equal(layout.rows, 1);
  assert.equal(layout.items.length, 3);
});

test('layoutPhotoCollage wraps many images instead of overflowing one page', () => {
  const sizes = Array.from({ length: 8 }, () => ({ width: 800, height: 600 }));
  const layout = layoutPhotoCollage(sizes, 520, 320, { gap: 4, frame: 2 });
  assert.equal(layout.items.length, 8);
  assert.ok(layout.rows >= 2);
  assert.ok(layout.width <= 520 + 0.01);
  assert.ok(layout.height <= 320 + 0.01);
});
