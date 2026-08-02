import type { SetupWeatherMetric } from '../types';

export const SETUP_WEATHER_METRIC_OPTIONS: Array<{ value: SetupWeatherMetric; label: string }> = [
  { value: 'temperature', label: 'Temperatur' },
  { value: 'temperature_min', label: 'Temperatur (min.)' },
  { value: 'temperature_max', label: 'Temperatur (max.)' },
  { value: 'condition', label: 'Wetterzustand' },
  { value: 'cloud_cover', label: 'Bewölkung' },
  { value: 'precipitation', label: 'Niederschlag' },
  { value: 'humidity', label: 'Luftfeuchtigkeit' },
  { value: 'wind_direction', label: 'Windrichtung' },
  { value: 'wind_speed', label: 'Windgeschwindigkeit' }
];

export function weatherMetricLabel(metric: SetupWeatherMetric | string | undefined): string {
  const match = SETUP_WEATHER_METRIC_OPTIONS.find((option) => option.value === metric);
  return match?.label || 'Temperatur';
}

export const DEFAULT_WEATHER_METRIC: SetupWeatherMetric = 'temperature';
