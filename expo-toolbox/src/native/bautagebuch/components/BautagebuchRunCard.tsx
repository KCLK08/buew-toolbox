import { ReactNode } from 'react';
import { Pressable, StyleSheet, Text, View } from 'react-native';
import { Swipeable } from 'react-native-gesture-handler';

import { PrimaryButton, StatusBadge } from '../../../components/mobile';
import { colors, spacing, typography } from '../../../constants/theme';
import type { BautagebuchRun } from '../types';

type Props = {
  run: BautagebuchRun;
  selected?: boolean;
  selectionMode?: boolean;
  onPress: () => void;
  onToggleSelect?: () => void;
  onRename: () => void;
  onDelete: () => void;
};

function formatShortUpdated(iso: string): string {
  const date = new Date(iso);
  const now = new Date();
  const time = date.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  if (date.toDateString() === now.toDateString()) {
    return `Heute, ${time}`;
  }
  return date.toLocaleString('de-DE', {
    day: '2-digit',
    month: '2-digit',
    year: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
}

function SwipeAction({
  label,
  tone,
  onPress
}: {
  label: string;
  tone: 'accent' | 'danger';
  onPress: () => void;
}) {
  return (
    <Pressable
      style={[styles.swipeAction, tone === 'danger' ? styles.swipeDanger : styles.swipeAccent]}
      onPress={onPress}
    >
      <Text style={styles.swipeLabel}>{label}</Text>
    </Pressable>
  );
}

export function BautagebuchRunCard({
  run,
  selected = false,
  selectionMode = false,
  onPress,
  onToggleSelect,
  onRename,
  onDelete
}: Props) {
  const renderRightActions = (): ReactNode => (
    <View style={styles.swipeRow}>
      <SwipeAction label="Umbenennen" tone="accent" onPress={onRename} />
      <SwipeAction label="Löschen" tone="danger" onPress={onDelete} />
    </View>
  );

  return (
    <Swipeable renderRightActions={selectionMode ? undefined : renderRightActions} overshootRight={false}>
      <Pressable onPress={selectionMode ? onToggleSelect : undefined}>
        <View style={[styles.card, selected ? styles.cardSelected : null]}>
          <View style={styles.titleBlock}>
            <Text style={styles.title} numberOfLines={2}>
              {run.title}
            </Text>
            <Text style={styles.updated}>{formatShortUpdated(run.updatedAt)}</Text>
          </View>

          {!selectionMode ? (
            <View style={styles.actionRow}>
              <StatusBadge
                label={run.status === 'completed' ? 'Abgeschlossen' : 'Entwurf'}
                tone={run.status === 'completed' ? 'success' : 'neutral'}
              />
              <PrimaryButton label="Öffnen" compact variant="secondary" onPress={onPress} />
            </View>
          ) : (
            <Text style={styles.hint}>{selected ? '✓ Ausgewählt' : 'Tippen zum Auswählen'}</Text>
          )}

          {!selectionMode ? (
            <Text style={styles.hint}>Nach links wischen für Umbenennen oder Löschen</Text>
          ) : null}
        </View>
      </Pressable>
    </Swipeable>
  );
}

const styles = StyleSheet.create({
  card: {
    gap: spacing.xs,
    paddingHorizontal: spacing.sm,
    paddingVertical: spacing.sm,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panel
  },
  cardSelected: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  titleBlock: {
    gap: 2
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink,
    fontSize: 15
  },
  updated: {
    ...typography.caption,
    color: colors.muted,
    fontSize: 12
  },
  actionRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  hint: {
    ...typography.caption,
    color: colors.muted,
    fontSize: 11
  },
  swipeRow: {
    flexDirection: 'row',
    marginBottom: spacing.xs
  },
  swipeAction: {
    width: 88,
    justifyContent: 'center',
    alignItems: 'center',
    marginLeft: 6,
    borderRadius: 10
  },
  swipeAccent: {
    backgroundColor: colors.accent
  },
  swipeDanger: {
    backgroundColor: colors.danger
  },
  swipeLabel: {
    ...typography.caption,
    color: colors.white,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 11
  }
});
