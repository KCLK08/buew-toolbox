import Constants from 'expo-constants';

type ExtraConfig = {
  supabaseUrl?: string;
  supabaseAnonKey?: string;
  toolboxWebBaseUrl?: string;
};

const extra = (Constants.expoConfig?.extra ?? {}) as ExtraConfig;

export const appConfig = {
  supabaseUrl: process.env.EXPO_PUBLIC_SUPABASE_URL || extra.supabaseUrl || '',
  supabaseAnonKey: process.env.EXPO_PUBLIC_SUPABASE_ANON_KEY || extra.supabaseAnonKey || '',
  toolboxWebBaseUrl:
    process.env.EXPO_PUBLIC_TOOLBOX_WEB_BASE_URL ||
    extra.toolboxWebBaseUrl ||
    'https://kclk08.github.io/buew-toolbox'
};

export function isSupabaseConfigured(): boolean {
  return Boolean(appConfig.supabaseUrl && appConfig.supabaseAnonKey);
}

export function toolWebUrl(webPath: string): string {
  const base = appConfig.toolboxWebBaseUrl.replace(/\/$/, '');
  const path = webPath.startsWith('/') ? webPath : `/${webPath}`;
  return `${base}${path}`;
}
