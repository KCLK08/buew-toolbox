const WEATHER_CATEGORY_KEYWORDS: Record<string, string[][]> = {
  clear: [['klar', 'sonnig', 'sonne', 'heiter', 'clear', 'sunny']],
  partly_cloudy: [
    ['teils bewölkt', 'teilweise bewolkt', 'leicht bewolkt', 'partly cloudy', 'wolkig'],
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
