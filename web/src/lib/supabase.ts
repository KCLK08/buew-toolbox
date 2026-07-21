import { createClient, type SupabaseClient } from '@supabase/supabase-js';
import type { Database } from '@buew/shared';

const url = import.meta.env.VITE_SUPABASE_URL as string;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY as string;

if (!url || !anonKey) {
  throw new Error('VITE_SUPABASE_URL und VITE_SUPABASE_ANON_KEY müssen gesetzt sein.');
}

export const supabase: SupabaseClient<Database> = createClient<Database>(url, anonKey, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
    storage: localStorage
  }
});

export const webConfig = {
  toolboxWebBaseUrl:
    (import.meta.env.VITE_TOOLBOX_WEB_BASE_URL as string | undefined) ??
    'https://kclk08.github.io/buew-toolbox',
  authRedirectUrl:
    (import.meta.env.VITE_AUTH_REDIRECT_URL as string | undefined) ??
    `${window.location.origin}${import.meta.env.BASE_URL}`
};
