import { Pressable, StyleSheet, Text, View } from 'react-native';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { Card } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { type ProjectListItem } from '../../lib/btb-filter';
import { formatRunCount } from '../../lib/group-runs-by-calendar';

type Props = {
  projects: ProjectListItem[];
  onSelectProject: (projectKey: string) => void;
};

export function ProjectFilterList({ projects, onSelectProject }: Props) {
  if (projects.length === 0) {
    return (
      <Card style={styles.emptyCard}>
        <Text style={styles.emptyText}>Keine Projekte vorhanden</Text>
      </Card>
    );
  }

  return (
    <View style={styles.root}>
      {projects.map((project) => (
        <Pressable
          key={project.projectKey}
          accessibilityRole="button"
          onPress={() => onSelectProject(project.projectKey)}
        >
          <Card style={styles.projectCard}>
            <View style={styles.iconWrap}>
              <MaterialCommunityIcons name="domain" size={22} color={colors.accent} />
            </View>
            <View style={styles.copy}>
              <Text style={styles.title} numberOfLines={2}>
                {project.projectLabel}
              </Text>
              <Text style={styles.meta}>{formatRunCount(project.runCount)}</Text>
            </View>
            <MaterialCommunityIcons name="chevron-right" size={22} color={colors.muted} />
          </Card>
        </Pressable>
      ))}
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    gap: spacing.sm
  },
  projectCard: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm,
    minHeight: spacing.touchMin + 8
  },
  iconWrap: {
    width: 44,
    height: 44,
    borderRadius: 12,
    backgroundColor: colors.badgeBg,
    alignItems: 'center',
    justifyContent: 'center'
  },
  copy: {
    flex: 1,
    gap: 2
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  meta: {
    ...typography.caption,
    color: colors.muted
  },
  emptyCard: {
    alignItems: 'center',
    paddingVertical: spacing.lg
  },
  emptyText: {
    ...typography.body,
    color: colors.muted
  }
});
