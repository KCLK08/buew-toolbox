import { ReactNode } from 'react';
import { Pressable, StyleSheet, Text, View } from 'react-native';
import { Swipeable } from 'react-native-gesture-handler';

import { Card, StatusBadge } from '../../../components/mobile';
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
  const dateLabel =
    run.title.match(/\d{4}-\d{2}-\d{2}/)?.[0] ||
    new Date(run.createdAt).toLocaleDateString('de-DE');

  const renderRightActions = (): ReactNode => (
    <View style={styles.swipeRow}>
      <SwipeAction label="Umbenennen" tone="accent" onPress={onRename} />
      <SwipeAction label="Löschen" tone="danger" onPress={onDelete} />
    </View>
  );

  return (
    <Swipeable renderRightActions={selectionMode ? undefined : renderRightActions} overshootRight={false}>
      <Pressable onPress={selectionMode ? onToggleSelect : onPress}>
        <Card style={selected ? { ...styles.card, ...styles.cardSelected } : styles.card}>
          <View style={styles.headerRow}>
            <View style={styles.titleBlock}>
              <Text style={styles.title} numberOfLines={2}>
                {run.title}
              </Text>
              <Text style={styles.meta}>{dateLabel}</Text>
            </View>
            <StatusBadge
              label={run.status === 'completed' ? 'Abgeschlossen' : 'Entwurf'}
              tone={run.status === 'completed' ? 'success' : 'neutral'}
            />
          </View>
          <Text style={styles.updated}>
            Zuletzt geändert/bearbeitet am:{' '}
            {new Date(run.updatedAt).toLocaleString('de-DE', {
              day: '2-digit',
              month: '2-digit',
              year: 'numeric',
              hour: '2-digit',
              minute: '2-digit'
            })}
          </Text>
          {!selectionMode ? (
            <Text style={styles.hint}>Nach links wischen für Aktionen · Tippen zum Öffnen</Text>
          ) : (
            <Text style={styles.hint}>{selected ? '✓ Ausgewählt' : 'Tippen zum Auswählen'}</Text>
          )}
        </Card>
      </Pressable>
    </Swipeable>
  );
}

const styles = StyleSheet.create({
  card: {
    gap: spacing.sm
  },
  cardSelected: {
    borderColor: colors.accent,
    backgroundColor: colors.badgeBg
  },
  headerRow: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  titleBlock: {
    flex: 1,
    gap: 4
  },
  title: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  meta: {
    ...typography.caption,
    color: colors.muted
  },
  updated: {
    ...typography.caption,
    color: colors.muted
  },
  hint: {
    ...typography.caption,
    color: colors.muted
  },
  swipeRow: {
    flexDirection: 'row',
    marginBottom: spacing.sm
  },
  swipeAction: {
    width: 96,
    justifyContent: 'center',
    alignItems: 'center',
    marginLeft: 8,
    borderRadius: 12
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
    fontFamily: 'SpaceGrotesk_600SemiBold'
  }
});
