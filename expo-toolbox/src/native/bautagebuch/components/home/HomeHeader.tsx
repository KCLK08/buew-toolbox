import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../../../constants/theme';

export function HomeHeader() {
  return (
    <View style={styles.root}>
      <Text style={styles.title}>Bautagebuch</Text>
      <Text style={styles.subtitle}>Digitale Baustellendokumentation schnell und strukturiert erfassen</Text>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.xxs,
    paddingBottom: spacing.xs
  },
  title: {
    ...typography.title,
    color: colors.ink,
    fontSize: 24,
    lineHeight: 30
  },
  subtitle: {
    ...typography.caption,
    color: colors.muted,
    fontSize: 14,
    lineHeight: 20
  }
});
