import { Platform } from 'react-native';
import type { BottomTabNavigationOptions } from '@react-navigation/bottom-tabs';

import { colors, spacing, typography } from '../constants/theme';

export function tabBarBottomInset(insets: { bottom: number }) {
  return Math.max(insets.bottom, Platform.OS === 'android' ? 20 : spacing.xs);
}

export function tabBarTotalHeight(insets: { bottom: number }) {
  return spacing.tabBarBody + tabBarBottomInset(insets);
}

export function createAppTabScreenOptions(insets: { bottom: number }): BottomTabNavigationOptions {
  const safeBottom = tabBarBottomInset(insets);
  return {
    headerShown: false,
    tabBarActiveTintColor: colors.tabActive,
    tabBarInactiveTintColor: colors.tabInactive,
    tabBarLabelStyle: {
      ...typography.caption,
      fontFamily: 'SpaceGrotesk_600SemiBold',
      fontSize: 11
    },
    tabBarStyle: {
      backgroundColor: colors.panel,
      borderTopColor: colors.border,
      borderTopWidth: 1,
      height: tabBarTotalHeight(insets),
      paddingBottom: safeBottom,
      paddingTop: spacing.xs
    },
    tabBarItemStyle: {
      minHeight: spacing.touchMin
    }
  };
}
