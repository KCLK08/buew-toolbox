import * as Haptics from 'expo-haptics';
import { Platform } from 'react-native';

function canHaptic() {
  return Platform.OS === 'ios' || Platform.OS === 'android';
}

export async function hapticLight() {
  if (!canHaptic()) return;
  await Haptics.impactAsync(Haptics.ImpactFeedbackStyle.Light);
}

export async function hapticMedium() {
  if (!canHaptic()) return;
  await Haptics.impactAsync(Haptics.ImpactFeedbackStyle.Medium);
}

export async function hapticSuccess() {
  if (!canHaptic()) return;
  await Haptics.notificationAsync(Haptics.NotificationFeedbackType.Success);
}

export async function hapticSelection() {
  if (!canHaptic()) return;
  await Haptics.selectionAsync();
}
