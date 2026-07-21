import 'react-native-url-polyfill/auto';

import Constants from 'expo-constants';
import { createClient, type SupabaseClient } from '@supabase/supabase-js';
import type { Database } from '@buew/shared';

import { secureAuthStorage } from './secureAuthStorage';

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

if (!appConfig.supabaseUrl || !appConfig.supabaseAnonKey) {
  throw new Error('EXPO_PUBLIC_SUPABASE_URL und EXPO_PUBLIC_SUPABASE_ANON_KEY müssen gesetzt sein.');
}

export const supabase: SupabaseClient<Database> = createClient<Database>(
  appConfig.supabaseUrl,
  appConfig.supabaseAnonKey,
  {
    auth: {
      storage: secureAuthStorage,
      autoRefreshToken: true,
      persistSession: true,
      detectSessionInUrl: false
    }
  }
);

export function toolWebUrl(webPath: string): string {
  const base = appConfig.toolboxWebBaseUrl.replace(/\/$/, '');
  const path = webPath.startsWith('/') ? webPath : `/${webPath}`;
  return `${base}${path}`;
}
