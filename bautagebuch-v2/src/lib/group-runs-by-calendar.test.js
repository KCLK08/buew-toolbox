import assert from 'node:assert/strict';

import { groupRunsByCalendar, resolveProjectLabel } from './group-runs-by-calendar.js';

function run(runId, title, updatedAt, values = {}) {
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

function test(name, fn) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('resolveProjectLabel reads project name from BTB title', () => {
  const label = resolveProjectLabel(run('r1', 'BTB_Tunnel_Süd_2025-06-12', '2025-06-12T10:00:00Z'), null);
  assert.equal(label, 'Tunnel Süd');
});

test('groupRunsByCalendar sorts same-day runs by updatedAt descending', () => {
  const tree = groupRunsByCalendar([
    run('older', 'BTB_Projekt_A_2025-06-12', '2025-06-12T08:00:00Z'),
    run('newer', 'BTB_Projekt_A_2025-06-12', '2025-06-12T18:00:00Z')
  ]);
  const week = tree.years.flatMap((year) => year.weeks)[0];
  const runsInWeek = week.projects.flatMap((project) => project.runs);
  assert.equal(runsInWeek[0].runId, 'newer');
});

console.log(`\n${passed} passed`);
