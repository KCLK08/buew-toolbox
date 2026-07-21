import * as Location from 'expo-location';

export async function syncWeatherValues(): Promise<{
  weather: string;
  tempMin: string;
  tempMax: string;
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
    '&current=temperature_2m,weather_code&daily=temperature_2m_max,temperature_2m_min&timezone=auto';

  const response = await fetch(url);
  if (!response.ok) {
    throw new Error('Wetterdaten konnten nicht geladen werden.');
  }

  const data = (await response.json()) as {
    current?: { temperature_2m?: number; weather_code?: number };
    daily?: { temperature_2m_max?: number[]; temperature_2m_min?: number[] };
  };

  const code = Number(data.current?.weather_code ?? 0);
  const weather = mapWeatherCode(code);
  const tempMin = String(data.daily?.temperature_2m_min?.[0] ?? data.current?.temperature_2m ?? '');
  const tempMax = String(data.daily?.temperature_2m_max?.[0] ?? data.current?.temperature_2m ?? '');

  return { weather, tempMin, tempMax };
}

function mapWeatherCode(code: number): string {
  if ([0].includes(code)) return 'sonnig';
  if ([1, 2, 3].includes(code)) return 'heiter bis wolkig';
  if ([45, 48].includes(code)) return 'Nebel';
  if ([51, 53, 55, 56, 57].includes(code)) return 'Nieselregen';
  if ([61, 63, 65, 66, 67, 80, 81, 82].includes(code)) return 'Regen';
  if ([71, 73, 75, 77, 85, 86].includes(code)) return 'Schnee';
  if ([95, 96, 99].includes(code)) return 'Gewitter';
  return 'wechselhaft';
}
