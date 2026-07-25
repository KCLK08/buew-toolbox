import assert from 'node:assert/strict';

import {
  buildBtbTitle,
  buildFotodokuTitle,
  fotodokuTitleFromBtbTitle,
  parseBtbTitle,
  sanitizeBtbInput
} from './btb-naming';

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('sanitizeBtbInput replaces spaces with underscores', () => {
  assert.equal(sanitizeBtbInput('Tunnel Süd'), 'Tunnel_Süd');
});

test('buildBtbTitle follows BTB_INPUT_Datum convention', () => {
  assert.equal(buildBtbTitle('Tunnel Süd', '2025-07-25'), 'BTB_Tunnel_Süd_2025-07-25');
});

test('buildFotodokuTitle follows BTB_Fotodoku_INPUT_Datum convention', () => {
  assert.equal(
    buildFotodokuTitle('Strecke Nord', '2025-07-25'),
    'BTB_Fotodoku_Strecke_Nord_2025-07-25'
  );
});

test('parseBtbTitle reads new BTB titles', () => {
  assert.deepEqual(parseBtbTitle('BTB_Tunnel_Süd_2025-06-12'), {
    input: 'Tunnel Süd',
    date: '2025-06-12',
    kind: 'btb'
  });
});

test('parseBtbTitle reads legacy BTB titles', () => {
  assert.deepEqual(parseBtbTitle('BTB 2025-06-12 - Tunnel Süd'), {
    input: 'Tunnel Süd',
    date: '2025-06-12',
    kind: 'btb'
  });
});

test('fotodokuTitleFromBtbTitle derives fotodoku export name', () => {
  assert.equal(
    fotodokuTitleFromBtbTitle('BTB_Tunnel_Süd_2025-06-12'),
    'BTB_Fotodoku_Tunnel_Süd_2025-06-12'
  );
});

console.log(`\n${passed} tests passed`);
