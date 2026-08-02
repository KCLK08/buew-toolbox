import * as Location from 'expo-location';

import { pickWeatherDropdownOption } from '../lib/weather-dropdown';
import type { WeatherSnapshot } from '../lib/weather-run';

const WIND_DIRECTIONS = ['N', 'NO', 'O', 'SO', 'S', 'SW', 'W', 'NW'];

function formatTemperatureValue(value: unknown): string {
  const number = Number(value);
  if (!Number.isFinite(number)) return '';
  return String(Math.round(number));
}

function formatPercentValue(value: unknown): string {
  const number = Number(value);
  if (!Number.isFinite(number)) return '';
  return `${Math.round(number)} %`;
}

function formatPrecipitationValue(value: unknown): string {
  const number = Number(value);
  if (!Number.isFinite(number)) return '';
  return `${number.toFixed(1).replace('.', ',')} mm`;
}

function formatWindSpeedValue(value: unknown): string {
  const number = Number(value);
  if (!Number.isFinite(number)) return '';
  return `${Math.round(number)} km/h`;
}

function formatWindDirectionValue(value: unknown): string {
  const degrees = Number(value);
  if (!Number.isFinite(degrees)) return '';
  const index = Math.round(degrees / 45) % WIND_DIRECTIONS.length;
  return WIND_DIRECTIONS[index] || 'N';
}

function mapWeatherCodeFallback(code: number): string {
  if ([0].includes(code)) return 'sonnig';
  if ([1, 2, 3].includes(code)) return 'heiter bis wolkig';
  if ([45, 48].includes(code)) return 'Nebel';
  if ([51, 53, 55, 56, 57].includes(code)) return 'Nieselregen';
  if ([61, 63, 65, 66, 67, 80, 81, 82].includes(code)) return 'Regen';
  if ([71, 73, 75, 77, 85, 86].includes(code)) return 'Schnee';
  if ([95, 96, 99].includes(code)) return 'Gewitter';
  return 'wechselhaft';
}

export async function fetchWeatherSnapshot(): Promise<WeatherSnapshot> {
  const permission = await Location.requestForegroundPermissionsAsync();
  if (!permission.granted) {
    throw new Error('Standortberechtigung ist für Wetterdaten erforderlich.');
  }

  const position = await Location.getCurrentPositionAsync({
    accuracy: Location.Accuracy.Balanced
  });
  const { latitude, longitude } = position.coords;
  const url =
    `https://api.open-meteo.com/v1/forecast?latitude=${latitude}&longitude=${longitude}` +
    '&current=weather_code,temperature_2m,relative_humidity_2m,precipitation,cloud_cover,wind_speed_10m,wind_direction_10m' +
    '&daily=temperature_2m_max,temperature_2m_min&forecast_days=1&timezone=auto';

  const response = await fetch(url);
  if (!response.ok) {
    throw new Error('Wetterdaten konnten nicht geladen werden.');
  }

  const data = (await response.json()) as {
    current?: {
      temperature_2m?: number;
      weather_code?: number;
      relative_humidity_2m?: number;
      precipitation?: number;
      cloud_cover?: number;
      wind_speed_10m?: number;
      wind_direction_10m?: number;
    };
    daily?: { temperature_2m_max?: number[]; temperature_2m_min?: number[] };
  };

  const weatherCode = Number(data.current?.weather_code ?? 0);
  const temperature = formatTemperatureValue(data.current?.temperature_2m);
  const temperatureMin = formatTemperatureValue(
    data.daily?.temperature_2m_min?.[0] ?? data.current?.temperature_2m
  );
  const temperatureMax = formatTemperatureValue(
    data.daily?.temperature_2m_max?.[0] ?? data.current?.temperature_2m
  );

  return {
    temperature,
    temperatureMin,
    temperatureMax,
    condition: mapWeatherCodeFallback(weatherCode),
    cloudCover: formatPercentValue(data.current?.cloud_cover),
    precipitation: formatPrecipitationValue(data.current?.precipitation),
    humidity: formatPercentValue(data.current?.relative_humidity_2m),
    windDirection: formatWindDirectionValue(data.current?.wind_direction_10m),
    windSpeed: formatWindSpeedValue(data.current?.wind_speed_10m),
    weatherCode
  };
}

export async function syncWeatherValues(dropdownOptions: string[] = []): Promise<{
  weather: string;
  tempMin: string;
  tempMax: string;
  weatherCode: number;
}> {
  const snapshot = await fetchWeatherSnapshot();
  const weather =
    pickWeatherDropdownOption(dropdownOptions, snapshot.weatherCode) || snapshot.condition;
  return {
    weather,
    tempMin: snapshot.temperatureMin,
    tempMax: snapshot.temperatureMax,
    weatherCode: snapshot.weatherCode
  };
}

export { pickWeatherDropdownOption } from '../lib/weather-dropdown';
