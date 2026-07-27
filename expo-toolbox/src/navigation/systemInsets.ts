import { Platform } from 'react-native';

import { spacing } from '../constants/theme';

/**
 * Fallback when Android edge-to-edge reports `insets.bottom === 0`
 * but the gesture / 3-button navigation bar is still visible.
 */
export const ANDROID_NAV_BAR_FALLBACK = 48;

/** Extra breathing room between app tab labels and the Android system navigation bar. */
export const ANDROID_TAB_BAR_NAV_GAP = 10;

export function systemBottomInset(insets: { bottom: number }): number {
  if (Platform.OS === 'android') {
    return Math.max(insets.bottom, ANDROID_NAV_BAR_FALLBACK);
  }
  return Math.max(insets.bottom, spacing.sm);
}

/** Bottom inset for tab screens: system nav bar + small gap above it for tab labels. */
export function tabBarBottomInset(insets: { bottom: number }): number {
  if (Platform.OS === 'android') {
    return systemBottomInset(insets) + ANDROID_TAB_BAR_NAV_GAP;
  }
  return systemBottomInset(insets);
}
