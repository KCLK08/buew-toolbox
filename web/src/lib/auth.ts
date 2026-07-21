import { createAuthService, createBrowserNetworkMonitor } from '@buew/shared';

import { supabase, webConfig } from './supabase';

export const authService = createAuthService({
  client: supabase,
  emailRedirectTo: webConfig.authRedirectUrl,
  passwordResetRedirectTo: webConfig.authRedirectUrl
});

export const networkMonitor = createBrowserNetworkMonitor();
