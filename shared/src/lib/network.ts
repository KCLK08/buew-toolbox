export type NetworkStatus = {
  online: boolean;
};

export type NetworkMonitor = {
  getStatus: () => Promise<NetworkStatus>;
  subscribe: (listener: (status: NetworkStatus) => void) => () => void;
};

export function createBrowserNetworkMonitor(): NetworkMonitor {
  return {
    async getStatus() {
      return { online: typeof navigator === 'undefined' ? true : navigator.onLine };
    },
    subscribe(listener) {
      if (typeof window === 'undefined') {
        return () => undefined;
      }

      const onOnline = () => listener({ online: true });
      const onOffline = () => listener({ online: false });
      window.addEventListener('online', onOnline);
      window.addEventListener('offline', onOffline);
      return () => {
        window.removeEventListener('online', onOnline);
        window.removeEventListener('offline', onOffline);
      };
    }
  };
}
