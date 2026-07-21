import { LinearGradient } from 'expo-linear-gradient';
import { StyleSheet, View } from 'react-native';
import type { PropsWithChildren } from 'react';

import { colors } from '../constants/theme';

export function ToolboxBackground({ children }: PropsWithChildren) {
  return (
    <View style={styles.root}>
      <LinearGradient
        colors={[colors.gradientTop, colors.gradientBottom]}
        style={StyleSheet.absoluteFill}
      />
      <LinearGradient
        colors={[colors.gradientWarm, 'transparent']}
        start={{ x: 0.1, y: 0 }}
        end={{ x: 0.55, y: 0.55 }}
        style={[StyleSheet.absoluteFill, styles.warmGlow]}
        pointerEvents="none"
      />
      <LinearGradient
        colors={[colors.gradientCool, 'transparent']}
        start={{ x: 0.95, y: 0 }}
        end={{ x: 0.35, y: 0.6 }}
        style={[StyleSheet.absoluteFill, styles.coolGlow]}
        pointerEvents="none"
      />
      {children}
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg
  },
  warmGlow: {
    opacity: 0.9
  },
  coolGlow: {
    opacity: 0.85
  }
});
