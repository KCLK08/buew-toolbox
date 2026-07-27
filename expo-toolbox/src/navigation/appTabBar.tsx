import type { BottomTabNavigationOptions } from '@react-navigation/bottom-tabs';

import { colors, spacing, typography } from '../constants/theme';
import { tabBarBottomInset } from './systemInsets';

/** Visible tab content area (icon + label), excluding safe-area padding. */
export const TAB_BAR_CONTENT_HEIGHT = 54;
export const TAB_BAR_TOP_PADDING = 10;

export { tabBarBottomInset } from './systemInsets';

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
      marginBottom: 0
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
