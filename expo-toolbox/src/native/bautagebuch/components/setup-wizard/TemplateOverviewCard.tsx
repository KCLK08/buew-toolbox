import { ReactNode } from 'react';
import { Pressable, StyleSheet, Text, View } from 'react-native';
import { Swipeable } from 'react-native-gesture-handler';
import { MaterialCommunityIcons } from '@expo/vector-icons';

import { PrimaryButton, SingleLineText, StatusBadge } from '../../../../components/mobile';
import { colors, shadows, spacing, typography } from '../../../../constants/theme';
import { templateDisplayStatus } from '../../lib/setup-mapping';
import type { BautagebuchTemplate } from '../../types';

type CardAction = {
  label: string;
  variant: 'primary' | 'secondary' | 'ghost';
  action: 'open' | 'activate' | 'continue' | 'view';
};

type Props = {
  template: BautagebuchTemplate;
  isActive: boolean;
  action: CardAction;
  onAction: () => void;
  onActivate?: () => void;
  onArchive?: () => void;
  onDelete?: () => void;
  canArchive?: boolean;
  canDelete?: boolean;
};

function SwipeAction({
  label,
  tone,
  onPress
}: {
  label: string;
  tone: 'neutral' | 'danger';
  onPress: () => void;
}) {
  return (
    <Pressable
      style={[styles.swipeAction, tone === 'danger' ? styles.swipeDanger : styles.swipeNeutral]}
      onPress={onPress}
    >
      <Text style={styles.swipeLabel}>{label}</Text>
    </Pressable>
  );
}

export function TemplateOverviewCard({
  template,
  isActive,
  action,
  onAction,
  onActivate,
  onArchive,
  onDelete,
  canArchive = false,
  canDelete = false
}: Props) {
  const status = templateDisplayStatus(template.status, isActive);
  const swipeActions: Array<{ label: string; tone: 'neutral' | 'danger'; onPress: () => void }> = [];

  if (canArchive && onArchive) {
    swipeActions.push({ label: 'Archivieren', tone: 'neutral', onPress: onArchive });
  }
  if (canDelete && onDelete) {
    swipeActions.push({ label: 'Löschen', tone: 'danger', onPress: onDelete });
  }

  const renderRightActions = (): ReactNode => {
    if (swipeActions.length === 0) return null;
    return (
      <View style={styles.swipeRow}>
        {swipeActions.map((entry) => (
          <SwipeAction key={entry.label} label={entry.label} tone={entry.tone} onPress={entry.onPress} />
        ))}
      </View>
    );
  };

  const card = (
    <View
      style={[styles.templateCard, isActive ? styles.templateCardActive : null, shadows.card]}
    >
      <View style={styles.cardTop}>
        <View style={styles.cardTitleBlock}>
          <SingleLineText style={styles.templateName}>{template.templateName}</SingleLineText>
          <SingleLineText style={styles.fileName}>{template.fileName}</SingleLineText>
        </View>
        <StatusBadge label={status.label} tone={status.tone} />
      </View>

      <View style={styles.metaRow}>
        <View style={styles.metaItem}>
          <MaterialCommunityIcons name="file-pdf-box" size={16} color={colors.muted} />
          <Text style={styles.metaText}>{template.pageCount} Seite(n)</Text>
        </View>
        {isActive ? (
          <View style={styles.metaItem}>
            <MaterialCommunityIcons name="check-circle" size={16} color={colors.success} />
            <Text style={[styles.metaText, styles.metaTextActive]}>Aktiv für neue BTBs</Text>
          </View>
        ) : null}
      </View>

      <View style={styles.actionRow}>
        <PrimaryButton label={action.label} variant={action.variant} onPress={onAction} />
        {template.status === 'ready' && !isActive && onActivate ? (
          <Pressable style={styles.linkBtn} onPress={onActivate}>
            <Text style={styles.linkLabel}>Aktivieren</Text>
          </Pressable>
        ) : null}
      </View>

      {swipeActions.length > 0 ? (
        <Text style={styles.hint}>Nach links wischen für Optionen</Text>
      ) : null}
    </View>
  );

  if (swipeActions.length === 0) {
    return card;
  }

  return (
    <Swipeable renderRightActions={renderRightActions} overshootRight={false}>
      {card}
    </Swipeable>
  );
}

const styles = StyleSheet.create({
  templateCard: {
    gap: spacing.sm,
    padding: spacing.cardPadding,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border
  },
  templateCardActive: {
    borderColor: colors.accent,
    backgroundColor: colors.panel
  },
  cardTop: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  cardTitleBlock: {
    flex: 1,
    gap: 2,
    minWidth: 0
  },
  templateName: {
    ...typography.bodyStrong,
    color: colors.ink,
    fontSize: 17
  },
  fileName: {
    ...typography.caption,
    color: colors.muted
  },
  metaRow: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.sm
  },
  metaItem: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.xxs
  },
  metaText: {
    ...typography.caption,
    color: colors.muted
  },
  metaTextActive: {
    color: colors.success,
    fontFamily: 'SpaceGrotesk_600SemiBold'
  },
  actionRow: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    alignItems: 'center',
    gap: spacing.sm,
    paddingTop: spacing.xxs
  },
  linkBtn: {
    paddingVertical: spacing.xxs
  },
  linkLabel: {
    ...typography.caption,
    color: colors.accent,
    fontFamily: 'SpaceGrotesk_600SemiBold'
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
    width: 92,
    justifyContent: 'center',
    alignItems: 'center',
    marginLeft: 6,
    borderRadius: 10,
    paddingHorizontal: spacing.xxs
  },
  swipeNeutral: {
    backgroundColor: colors.accent
  },
  swipeDanger: {
    backgroundColor: colors.danger
  },
  swipeLabel: {
    ...typography.caption,
    color: colors.white,
    fontFamily: 'SpaceGrotesk_600SemiBold',
    fontSize: 11,
    textAlign: 'center'
  }
});
