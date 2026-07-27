import { Platform } from 'react-native';
import type { BottomTabNavigationOptions } from '@react-navigation/bottom-tabs';

import { colors, spacing, typography } from '../constants/theme';

/** Visible tab content area (icon + label), excluding safe-area padding. */
export const TAB_BAR_CONTENT_HEIGHT = 54;
export const TAB_BAR_TOP_PADDING = 10;

export function tabBarBottomInset(insets: { bottom: number }) {
  if (Platform.OS === 'android') {
    return Math.max(insets.bottom, 28);
  }
  return Math.max(insets.bottom, spacing.sm);
}

export function tabBarTotalHeight(insets: { bottom: number }) {
  return TAB_BAR_CONTENT_HEIGHT + TAB_BAR_TOP_PADDING + tabBarBottomInset(insets);
}

/** Height to reserve in scroll content above the tab bar. */
export function tabBarReservedHeight(insets: { bottom: number }) {
  return tabBarTotalHeight(insets);
}

export function createAppTabScreenOptions(insets: { bottom: number }): BottomTabNavigationOptions {
  const safeBottom = tabBarBottomInset(insets);
  const height = tabBarTotalHeight(insets);

  return {
    headerShown: false,
    tabBarActiveTintColor: colors.tabActive,
    tabBarInactiveTintColor: colors.tabInactive,
    tabBarLabelPosition: 'below-icon',
    tabBarAllowFontScaling: false,
    tabBarLabelStyle: {
      ...typography.caption,
      fontFamily: 'SpaceGrotesk_600SemiBold',
      fontSize: 11,
      lineHeight: 14,
      marginTop: 2,
      marginBottom: 2
    },
    tabBarIconStyle: {
      marginBottom: 0
    },
    tabBarItemStyle: {
      paddingVertical: 4,
      justifyContent: 'center'
    },
    tabBarStyle: {
      backgroundColor: colors.panel,
      borderTopColor: colors.border,
      borderTopWidth: 1,
      height,
      paddingTop: TAB_BAR_TOP_PADDING,
      paddingBottom: safeBottom
    }
  };
}
