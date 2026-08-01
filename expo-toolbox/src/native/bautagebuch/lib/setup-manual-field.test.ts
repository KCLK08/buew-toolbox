import assert from 'node:assert/strict';
import { test } from 'node:test';

import {
  geometryDraftFromField,
  isManualMappingField,
  mappingFieldGeometry
} from './setup-manual-field';
import type { MappingField } from './setup-mapping';

const manualField: MappingField = {
  fieldId: 'm1',
  fieldName: 'm1',
  labelCandidate: 'Test',
  type: 'text',
  options: [],
  page: 2,
  orderIndex: 0,
  displayOrder: 1,
  rect: [10, 20, 50, 60],
  geometry: {
    page: 2,
    rect: { x: 10, y: 20, width: 40, height: 40 }
  },
  source: 'manual'
};

test('isManualMappingField detects manual source', () => {
  assert.equal(isManualMappingField(manualField), true);
  assert.equal(isManualMappingField({ ...manualField, source: 'acroform' }), false);
});

test('geometryDraftFromField prefers stored geometry', () => {
  const draft = geometryDraftFromField(manualField);
  assert.deepEqual(draft, {
    page: 2,
    rect: { x: 10, y: 20, width: 40, height: 40 }
  });
});

test('mappingFieldGeometry falls back to legacy rect', () => {
  const geometry = mappingFieldGeometry({
    ...manualField,
    geometry: null,
    rect: [0, 0, 30, 20]
  });
  assert.deepEqual(geometry, {
    page: 2,
    rect: { x: 0, y: 0, width: 30, height: 20 }
  });
});

console.log('setup-manual-field tests passed');
