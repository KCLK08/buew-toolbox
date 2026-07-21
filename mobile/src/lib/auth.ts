import NetInfo from '@react-native-community/netinfo';
import { createAuthService, type NetworkMonitor } from '@buew/shared';

import { supabase } from './supabase';

export const authService = createAuthService({
  client: supabase
});

export const networkMonitor: NetworkMonitor = {
  async getStatus() {
    const state = await NetInfo.fetch();
    return { online: Boolean(state.isConnected && state.isInternetReachable !== false) };
  },
  subscribe(listener) {
    return NetInfo.addEventListener((state) => {
      listener({
        online: Boolean(state.isConnected && state.isInternetReachable !== false)
      });
    });
  }
};
