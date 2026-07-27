import { useEffect } from 'react';
import { BackHandler, Platform } from 'react-native';
import { useRouter } from 'expo-router';

/**
 * Ensures Android hardware back follows navigation history and exits at the root.
 * Tab navigators use backBehavior="none" so back pops the parent stack instead of switching tabs.
 */
export function AndroidHardwareBack() {
  const router = useRouter();

  useEffect(() => {
    if (Platform.OS !== 'android') return undefined;

    const subscription = BackHandler.addEventListener('hardwareBackPress', () => {
      if (router.canGoBack()) {
        router.back();
        return true;
      }
      return false;
    });

    return () => subscription.remove();
  }, [router]);

  return null;
}
