import * as Location from 'expo-location';

const WEATHER_CATEGORY_KEYWORDS: Record<string, string[][]> = {
  clear: [['klar', 'sonnig', 'sonne', 'heiter', 'clear', 'sunny']],
  partly_cloudy: [
    ['teils bewolkt', 'teilweise bewolkt', 'leicht bewolkt', 'partly cloudy', 'wolkig'],
    ['bewolkt', 'heiter', 'klar']
  ],
  cloudy: [['bedeckt', 'stark bewolkt', 'overcast', 'cloudy', 'bewolkt']],
  fog: [['nebel', 'fog', 'mist']],
  rain: [['regen', 'regnerisch', 'schauer', 'niesel', 'drizzle', 'rain']],
  snow: [['schnee', 'schneefall', 'graupel', 'snow', 'sleet']],
  thunder: [['gewitter', 'thunder', 'sturm']]
};

function normalizeSearchText(value: string): string {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function weatherCategoryFromCode(code: number): string {
  const value = Number(code);
  if (value === 95 || value === 96 || value === 99) return 'thunder';
  if ((value >= 71 && value <= 77) || value === 85 || value === 86) return 'snow';
  if ((value >= 51 && value <= 67) || (value >= 80 && value <= 82)) return 'rain';
  if (value === 45 || value === 48) return 'fog';
  if (value === 0) return 'clear';
  if (value === 1 || value === 2 || value === 3) return 'partly_cloudy';
  return 'cloudy';
}

export function pickWeatherDropdownOption(options: string[] = [], weatherCode: number): string {
  const normalizedOptions = options
    .map((option) => ({
      raw: String(option || '').trim(),
      normalized: normalizeSearchText(option)
    }))
    .filter((entry) => entry.raw.length > 0 && entry.normalized.length > 0);
  if (normalizedOptions.length === 0) return '';

  const category = weatherCategoryFromCode(weatherCode);
  const keywordGroups = WEATHER_CATEGORY_KEYWORDS[category] || [];

  for (const keywords of keywordGroups) {
    let best: { option: string; score: number } | null = null;
    for (const option of normalizedOptions) {
      let score = 0;
      for (const keyword of keywords) {
        const normalizedKeyword = normalizeSearchText(keyword);
        if (!normalizedKeyword) continue;
        if (!option.normalized.includes(normalizedKeyword)) continue;
        score = Math.max(score, normalizedKeyword.length);
      }
      if (score === 0) continue;
      if (!best || score > best.score) {
        best = { option: option.raw, score };
      }
    }
    if (best?.option) return best.option;
  }

  return '';
}

function formatTemperatureValue(value: unknown): string {
  const number = Number(value);
  if (!Number.isFinite(number)) return '';
  return String(Math.round(number));
}

export async function syncWeatherValues(dropdownOptions: string[] = []): Promise<{
  weather: string;
  tempMin: string;
  tempMax: string;
  weatherCode: number;
}> {
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
    '&current=weather_code,temperature_2m&daily=weather_code,temperature_2m_max,temperature_2m_min&forecast_days=1&timezone=auto';

  const response = await fetch(url);
  if (!response.ok) {
    throw new Error('Wetterdaten konnten nicht geladen werden.');
  }

  const data = (await response.json()) as {
    current?: { temperature_2m?: number; weather_code?: number };
    daily?: { temperature_2m_max?: number[]; temperature_2m_min?: number[] };
  };

  const weatherCode = Number(data.current?.weather_code ?? 0);
  const picked =
    pickWeatherDropdownOption(dropdownOptions, weatherCode) ||
    mapWeatherCodeFallback(weatherCode);
  const tempMin = formatTemperatureValue(
    data.daily?.temperature_2m_min?.[0] ?? data.current?.temperature_2m
  );
  const tempMax = formatTemperatureValue(
    data.daily?.temperature_2m_max?.[0] ?? data.current?.temperature_2m
  );

  return { weather: picked, tempMin, tempMax, weatherCode };
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
