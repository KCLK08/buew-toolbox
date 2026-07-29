import assert from 'node:assert/strict';

import { mergeScannedFields } from './field-merge';
import {
  createManualFieldInput,
  fieldHasGeometry,
  fieldRectToLegacyRect,
  legacyRectToFieldRect,
  normalizeFieldGeometry,
  scanResultToTemplateFieldInput
} from './template-field';
import type { DetectedField } from '../types';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('legacyRectToFieldRect converts pdf boxes', () => {
  const rect = legacyRectToFieldRect([100, 200, 250, 260]);
  assert.equal(rect.x, 100);
  assert.equal(rect.y, 200);
  assert.equal(rect.width, 150);
  assert.equal(rect.height, 60);
  assert.deepEqual(fieldRectToLegacyRect(rect), [100, 200, 250, 260]);
});

test('createManualFieldInput stores manual source and geometry', () => {
  const field = createManualFieldInput({
    name: 'Unterschrift',
    type: 'signature',
    page: 2,
    rect: { x: 10, y: 20, width: 80, height: 30 }
  });
  assert.equal(field.source, 'manual');
  assert.equal(field.geometry?.page, 2);
  assert.equal(fieldHasGeometry({ geometry: field.geometry, rect: null }), true);
});

test('mergeScannedFields preserves manual fields', () => {
  const existing: DetectedField[] = [
    {
      id: '1',
      templateId: 'tpl',
      fieldId: 'manual_1',
      fieldName: 'manual_1',
      labelCandidate: 'Unterschrift',
      type: 'signature',
      options: [],
      page: 1,
      orderIndex: 0,
      rect: [10, 10, 50, 40],
      geometry: { page: 1, rect: { x: 10, y: 10, width: 40, height: 30 } },
      source: 'manual',
      createdAt: '',
      updatedAt: ''
    }
  ];
  const merged = mergeScannedFields(existing, [
    {
      fieldId: 'acro_1',
      fieldName: 'Baustelle',
      labelCandidate: 'Baustelle',
      type: 'text',
      options: [],
      page: 1,
      orderIndex: 1,
      rect: [100, 100, 200, 130]
    }
  ]);
  assert.equal(merged.length, 2);
  assert.equal(merged.filter((entry) => entry.source === 'manual').length, 1);
});

test('scanResultToTemplateFieldInput marks acroform geometry', () => {
  const input = scanResultToTemplateFieldInput(
    {
      fieldId: 'f1',
      fieldName: 'Datum',
      labelCandidate: 'Datum',
      type: 'datetime',
      options: [],
      page: 1,
      orderIndex: 0,
      rect: [1, 2, 3, 4]
    },
    'acroform'
  );
  assert.equal(input.source, 'acroform');
  assert.ok(normalizeFieldGeometry({ geometry: input.geometry, rect: input.rect }));
});

console.log(`\n${passed} template-field tests passed`);
