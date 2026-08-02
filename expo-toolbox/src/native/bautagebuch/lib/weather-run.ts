import { pickWeatherDropdownOption } from './weather-dropdown';
import type { SetupFieldConfig, SetupWeatherMetric } from '../types';
import { DEFAULT_WEATHER_METRIC } from './weather-metrics';

export type WeatherFieldRef = {
  fieldId: string;
  fieldName?: string;
  metric: SetupWeatherMetric;
  options?: string[];
};

export type WeatherSnapshot = {
  temperature: string;
  temperatureMin: string;
  temperatureMax: string;
  condition: string;
  cloudCover: string;
  precipitation: string;
  humidity: string;
  windDirection: string;
  windSpeed: string;
  weatherCode: number;
};

const LEGACY_WEATHER_FIELD_METRICS: Record<string, SetupWeatherMetric> = {
  Dropdown6: 'condition',
  Text11: 'temperature_min',
  Text12: 'temperature_max'
};

function normalizeMetric(value: string | undefined): SetupWeatherMetric {
  const metric = String(value || '').trim() as SetupWeatherMetric;
  if (
    metric === 'temperature' ||
    metric === 'temperature_min' ||
    metric === 'temperature_max' ||
    metric === 'condition' ||
    metric === 'cloud_cover' ||
    metric === 'precipitation' ||
    metric === 'humidity' ||
    metric === 'wind_direction' ||
    metric === 'wind_speed'
  ) {
    return metric;
  }
  return DEFAULT_WEATHER_METRIC;
}

export function isWeatherField(field: { type?: string; skipped?: boolean } | null | undefined): boolean {
  if (!field || field.skipped === true) return false;
  return String(field.type || '').trim() === 'weather';
}

export function resolveWeatherMetricForField(field: SetupFieldConfig): SetupWeatherMetric {
  if (field.weatherMetric) return normalizeMetric(field.weatherMetric);
  const legacy = LEGACY_WEATHER_FIELD_METRICS[String(field.fieldName || '').trim()];
  if (legacy) return legacy;
  return DEFAULT_WEATHER_METRIC;
}

export function listWeatherFieldsInSection(
  fields: Array<SetupFieldConfig | null | undefined>
): WeatherFieldRef[] {
  const refs: WeatherFieldRef[] = [];
  for (const field of fields) {
    if (!field || field.skipped === true) continue;
    const fieldId = String(field.fieldId || '').trim();
    if (!fieldId) continue;
    if (isWeatherField(field)) {
      refs.push({
        fieldId,
        fieldName: field.fieldName,
        metric: resolveWeatherMetricForField(field),
        options: Array.isArray(field.options) ? field.options : []
      });
      continue;
    }
    const legacyMetric = LEGACY_WEATHER_FIELD_METRICS[String(field.fieldName || '').trim()];
    if (legacyMetric) {
      refs.push({
        fieldId,
        fieldName: field.fieldName,
        metric: legacyMetric,
        options: Array.isArray(field.options) ? field.options : []
      });
    }
  }
  return refs;
}

export function sectionHasWeatherFields(
  fields: Array<SetupFieldConfig | null | undefined>
): boolean {
  return listWeatherFieldsInSection(fields).length > 0;
}

function formatCondition(snapshot: WeatherSnapshot, options: string[] = []): string {
  const picked = pickWeatherDropdownOption(options, snapshot.weatherCode);
  return picked || snapshot.condition;
}

export function resolveWeatherValueForMetric(
  metric: SetupWeatherMetric,
  snapshot: WeatherSnapshot,
  options: string[] = []
): string {
  switch (metric) {
    case 'temperature':
      return snapshot.temperature;
    case 'temperature_min':
      return snapshot.temperatureMin;
    case 'temperature_max':
      return snapshot.temperatureMax;
    case 'condition':
      return formatCondition(snapshot, options);
    case 'cloud_cover':
      return snapshot.cloudCover;
    case 'precipitation':
      return snapshot.precipitation;
    case 'humidity':
      return snapshot.humidity;
    case 'wind_direction':
      return snapshot.windDirection;
    case 'wind_speed':
      return snapshot.windSpeed;
    default:
      return snapshot.temperature;
  }
}

export function applyWeatherSnapshotToFields(
  weatherFields: WeatherFieldRef[],
  snapshot: WeatherSnapshot,
  values: Record<string, unknown>
): Record<string, unknown> {
  const nextValues = { ...values };
  for (const field of weatherFields) {
    const value = resolveWeatherValueForMetric(field.metric, snapshot, field.options);
    if (!value) continue;
    nextValues[`field:${field.fieldId}`] = value;
  }
  return nextValues;
}
