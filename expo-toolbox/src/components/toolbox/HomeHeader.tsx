import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';

export function HomeHeader() {
  return (
    <View style={styles.root}>
      <Text style={styles.title}>BÜW Toolbox</Text>
      <Text style={styles.subtitle}>Digitale Werkzeuge für eine strukturierte Baustellendokumentation</Text>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.xs,
    paddingBottom: spacing.sm
  },
  title: {
    ...typography.display,
    color: colors.ink,
    fontSize: 32,
    letterSpacing: -0.5
  },
  subtitle: {
    ...typography.body,
    color: colors.muted,
    fontSize: 16,
    lineHeight: 24,
    maxWidth: 520
  }
});
