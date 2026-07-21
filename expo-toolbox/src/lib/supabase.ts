import 'react-native-url-polyfill/auto';

import AsyncStorage from '@react-native-async-storage/async-storage';
import { createClient, type SupabaseClient } from '@supabase/supabase-js';

import { appConfig, isSupabaseConfigured } from './config';
import type { Database } from './database.types';

let client: SupabaseClient<Database> | null = null;

/**
 * Lazily creates the Supabase client once URL + anon key are configured.
 * Returns null until EXPO_PUBLIC_SUPABASE_URL / EXPO_PUBLIC_SUPABASE_ANON_KEY are set.
 */
export function getSupabase(): SupabaseClient<Database> | null {
  if (!isSupabaseConfigured()) {
    return null;
  }

  if (!client) {
    client = createClient<Database>(appConfig.supabaseUrl, appConfig.supabaseAnonKey, {
      auth: {
        storage: AsyncStorage,
        autoRefreshToken: true,
        persistSession: true,
        detectSessionInUrl: false
      }
    });
  }

  return client;
}

export type { Database };
