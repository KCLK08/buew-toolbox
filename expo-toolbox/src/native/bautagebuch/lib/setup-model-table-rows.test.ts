import assert from 'node:assert/strict';

import { buildRunSections } from './setup-model.js';
import { configuredTableRowCount, visibleRowCountForSection } from './run-validation';
import { buildTableSectionsFromWizard } from './setup-wizard-tables';
import { getWizardState } from './setup-mapping';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

function customTableModel(rowCount: number) {
  const wizard = getWizardState({
    wizard: {
      tables: [
        {
          tableId: 'tbl_workforce',
          label: 'Arbeitskräfte',
          rowCount,
          columns: [
            { columnId: 'c1', label: 'Mitarbeiter', type: 'text' },
            { columnId: 'c2', label: 'Stunden', type: 'text' },
            { columnId: 'c3', label: 'Beginn', type: 'text' },
            { columnId: 'c4', label: 'Ende', type: 'text' }
          ]
        }
      ],
      tableAssignments: {
        f1: { tableId: 'tbl_workforce', rowIndex: 0, columnId: 'c1' }
      }
    }
  });
  const fields = [
    {
      fieldId: 'f1',
      fieldName: 'Text1',
      labelCandidate: 'Mitarbeiter',
      type: 'text',
      page: 1,
      rect: [0, 0, 10, 10],
      source: 'detected',
      displayOrder: 1
    }
  ];
  const tableSections = buildTableSectionsFromWizard(wizard, fields as never);
  return {
    table_sections: tableSections,
    section_order: [{ kind: 'table', id: 'tbl_workforce' }]
  };
}

test('buildRunSections keeps configured empty table rows', () => {
  const model = customTableModel(5);
  const configuredTable = model.table_sections[0] as { rows: unknown[] };
  assert.equal(configuredTable.rows.length, 5);
  const sections = buildRunSections(model);
  assert.equal(sections.length, 1);
  const runTable = sections[0];
  assert.equal(runTable?.kind, 'table');
  if (runTable?.kind === 'table') {
    assert.equal(runTable.rows.length, 5);
  }
});

test('visibleRowCountForSection starts with one row for custom tables', () => {
  const model = customTableModel(5);
  const section = buildRunSections(model)[0] as never;
  const visible = visibleRowCountForSection(section, {}, {
    maxRows: configuredTableRowCount(model, 'tbl_workforce')
  });
  assert.equal(visible, 1);
});

test('visibleRowCountForSection respects stored row count up to configured max', () => {
  const model = customTableModel(5);
  const section = buildRunSections(model)[0] as never;
  const options = { maxRows: configuredTableRowCount(model, 'tbl_workforce') };
  assert.equal(visibleRowCountForSection(section, { '__tableRows:tbl_workforce': 3 }, options), 3);
  assert.equal(visibleRowCountForSection(section, { '__tableRows:tbl_workforce': 9 }, options), 5);
});

test('visibleRowCountForSection never exceeds configured table rows', () => {
  const model = customTableModel(4);
  const section = buildRunSections(model)[0] as never;
  const options = { maxRows: configuredTableRowCount(model, 'tbl_workforce') };
  assert.equal(visibleRowCountForSection(section, { '__tableRows:tbl_workforce': 99 }, options), 4);
});

console.log(`\n${passed} tests passed`);
