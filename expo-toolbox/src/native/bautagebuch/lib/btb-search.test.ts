import assert from 'node:assert/strict';

import { filterRunsBySearchQuery, runMatchesQuery } from './btb-search';
import type { BautagebuchRun } from '../types';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

function sampleRun(overrides: Partial<BautagebuchRun> = {}): BautagebuchRun {
  return {
    runId: 'run-1',
    templateId: 'tpl-1',
    title: 'BTB_Tunnel_Süd_2025-07-25',
    setupVersion: 1,
    values: { field_notes: 'Zaun wurde gesetzt' },
    sectionIndex: 0,
    status: 'draft',
    photoDoc: { enabled: null, entries: [], updatedAt: '2025-07-25T10:00:00.000Z' },
    createdAt: '2025-07-25T10:00:00.000Z',
    updatedAt: '2025-07-25T10:00:00.000Z',
    completedAt: '',
    deleted_at: null,
    ...overrides
  };
}

test('runMatchesQuery finds terms inside field values', () => {
  assert.equal(runMatchesQuery(sampleRun(), 'Zaun'), true);
});

test('runMatchesQuery is case-insensitive', () => {
  assert.equal(runMatchesQuery(sampleRun(), 'zaun'), true);
});

test('runMatchesQuery matches title text', () => {
  assert.equal(runMatchesQuery(sampleRun({ values: {} }), 'Tunnel'), true);
});

test('runMatchesQuery returns true for empty query', () => {
  assert.equal(runMatchesQuery(sampleRun(), ''), true);
  assert.equal(runMatchesQuery(sampleRun(), '   '), true);
});

test('runMatchesQuery returns false when term is missing', () => {
  assert.equal(runMatchesQuery(sampleRun(), 'Brücke'), false);
});

test('filterRunsBySearchQuery keeps only matching runs', () => {
  const runs = [
    sampleRun({ runId: 'a', values: { field_notes: 'Zaun montiert' } }),
    sampleRun({ runId: 'b', values: { field_notes: 'Asphalt verlegt' } })
  ];
  const filtered = filterRunsBySearchQuery(runs, 'Zaun');
  assert.equal(filtered.length, 1);
  assert.equal(filtered[0]?.runId, 'a');
});

console.log(`\n${passed} tests passed`);
