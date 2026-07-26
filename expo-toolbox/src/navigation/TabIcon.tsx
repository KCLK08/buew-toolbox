import { MaterialCommunityIcons } from '@expo/vector-icons';
import type { ComponentProps } from 'react';

import { colors } from '../constants/theme';

type IconName = ComponentProps<typeof MaterialCommunityIcons>['name'];

export function TabIcon({ name, focused }: { name: IconName; focused: boolean }) {
  return (
    <MaterialCommunityIcons
      name={name}
      size={22}
      color={focused ? colors.tabActive : colors.tabInactive}
    />
  );
}
