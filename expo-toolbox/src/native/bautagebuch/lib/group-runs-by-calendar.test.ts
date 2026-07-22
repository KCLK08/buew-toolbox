import assert from 'node:assert/strict';

import {
  groupRunsByCalendar,
  resolveProjectLabel
} from './group-runs-by-calendar';
import type { BautagebuchRun } from '../types';

function run(
  runId: string,
  title: string,
  updatedAt: string,
  values: Record<string, unknown> = {}
): BautagebuchRun {
  return {
    runId,
    templateId: 'tpl_test',
    title,
    setupVersion: 1,
    values,
    sectionIndex: 0,
    status: 'draft',
    photoDoc: { enabled: null, entries: [], updatedAt: '' },
    createdAt: updatedAt,
    updatedAt,
    completedAt: '',
    deleted_at: null
  };
}

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('resolveProjectLabel reads project name from BTB title', () => {
  const label = resolveProjectLabel(
    run('r1', 'BTB 2025-06-12 - Tunnel Süd', '2025-06-12T10:00:00Z'),
    null
  );
  assert.equal(label, 'Tunnel Süd');
});

test('groupRunsByCalendar sorts same-day runs by updatedAt descending', () => {
  const tree = groupRunsByCalendar([
    run('older', 'BTB 2025-06-12 - Projekt A', '2025-06-12T08:00:00Z'),
    run('newer', 'BTB 2025-06-12 - Projekt A', '2025-06-12T18:00:00Z')
  ]);

  const projectRuns = tree.years[0]?.weeks[0]?.projects[0]?.runs || [];
  assert.equal(projectRuns[0]?.runId, 'newer');
  assert.equal(projectRuns[1]?.runId, 'older');
});

test('groupRunsByCalendar groups runs by calendar week and project', () => {
  const tree = groupRunsByCalendar([
    run('r1', 'BTB 2025-06-12 - Nord', '2025-06-12T10:00:00Z'),
    run('r2', 'BTB 2025-06-13 - Süd', '2025-06-13T10:00:00Z')
  ]);

  const week = tree.years[0]?.weeks[0];
  assert.ok(week?.weekLabel.startsWith('KW '));
  assert.equal(week?.projects.length, 2);
  assert.equal(week?.runCount, 2);
});

console.log(`\n${passed} tests passed`);
