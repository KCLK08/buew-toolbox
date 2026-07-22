import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../../../constants/theme';

type Props = {
  issues: string[];
  onSelectIssue?: (index: number) => void;
};

export function SetupValidationList({ issues, onSelectIssue }: Props) {
  if (issues.length === 0) {
    return (
      <View style={styles.empty}>
        <Text style={styles.ok}>Alle Pflichtangaben sind vollständig.</Text>
      </View>
    );
  }

  return (
    <View style={styles.root}>
      <Text style={styles.title}>Bitte korrigieren:</Text>
      {issues.map((issue, index) => (
        <Pressable
          key={`${issue}-${index}`}
          style={styles.item}
          onPress={() => onSelectIssue?.(index)}
        >
          <Text style={styles.bullet}>•</Text>
          <Text style={styles.issue}>{issue}</Text>
        </Pressable>
      ))}
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.xs,
    padding: spacing.pageX
  },
  title: {
    ...typography.bodyStrong,
    color: colors.danger
  },
  item: {
    flexDirection: 'row',
    gap: spacing.xs,
    paddingVertical: spacing.xs
  },
  bullet: {
    ...typography.body,
    color: colors.danger
  },
  issue: {
    ...typography.body,
    color: colors.ink,
    flex: 1
  },
  empty: {
    padding: spacing.pageX
  },
  ok: {
    ...typography.body,
    color: colors.success
  }
});
