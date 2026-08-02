import assert from 'node:assert/strict';

import {
  applyWeatherSnapshotToFields,
  listWeatherFieldsInSection,
  resolveWeatherValueForMetric,
  sectionHasWeatherFields
} from './weather-run';
import type { WeatherSnapshot } from './weather-run';

const snapshot: WeatherSnapshot = {
  temperature: '18',
  temperatureMin: '12',
  temperatureMax: '22',
  condition: 'heiter',
  cloudCover: '35 %',
  precipitation: '0 mm',
  humidity: '62 %',
  windDirection: 'SW',
  windSpeed: '14 km/h',
  weatherCode: 2
};

let passed = 0;

function test(name: string, fn: () => void) {
  fn();
  passed += 1;
  console.log(`ok - ${name}`);
}

test('sectionHasWeatherFields detects weather type fields', () => {
  assert.equal(
    sectionHasWeatherFields([
      { fieldId: 'a', type: 'text' },
      { fieldId: 'b', type: 'weather', weatherMetric: 'humidity' }
    ]),
    true
  );
  assert.equal(sectionHasWeatherFields([{ fieldId: 'a', type: 'text' }]), false);
});

test('listWeatherFieldsInSection includes legacy ETB weather field names', () => {
  const fields = listWeatherFieldsInSection([
    { fieldId: 'd6', fieldName: 'Dropdown6', type: 'dropdown', options: ['Sonne', 'Regen'] },
    { fieldId: 't11', fieldName: 'Text11', type: 'text' },
    { fieldId: 'w1', type: 'weather', weatherMetric: 'wind_speed' }
  ]);
  assert.equal(fields.some((field) => field.fieldId === 'd6' && field.metric === 'condition'), true);
  assert.equal(fields.some((field) => field.fieldId === 't11' && field.metric === 'temperature_min'), true);
});

test('applyWeatherSnapshotToFields fills only configured weather fields', () => {
  const next = applyWeatherSnapshotToFields(
    [
      { fieldId: 'f1', metric: 'temperature' },
      { fieldId: 'f2', metric: 'wind_speed' }
    ],
    snapshot,
    {}
  );
  assert.equal(next['field:f1'], '18');
  assert.equal(next['field:f2'], '14 km/h');
});

test('resolveWeatherValueForMetric returns metric-specific values', () => {
  assert.equal(resolveWeatherValueForMetric('humidity', snapshot), '62 %');
  assert.equal(resolveWeatherValueForMetric('temperature_max', snapshot), '22');
});

console.log(`\n${passed} tests passed`);
