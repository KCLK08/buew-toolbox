import { Pressable, StyleSheet, Text, View } from 'react-native';

import { colors, shadows, spacing, typography } from '../../constants/theme';
import type { SiteReportExport } from '../../native/sitereport/db/database';
import { Card } from '../mobile';
import { hapticSelection } from '../../lib/haptics';
import { SelectionCheckbox } from './SelectionCheckbox';

type Props = {
  item: SiteReportExport;
  selected?: boolean;
  onSelectToggle?: () => void;
  onSharePdf?: () => void;
  onShareXlsx?: () => void;
  onDelete?: () => void;
  sharing?: boolean;
};

function FormatIcon({ format }: { format: 'pdf' | 'xlsx' }) {
  return (
    <View style={[styles.formatIcon, format === 'pdf' ? styles.pdfIcon : styles.xlsxIcon]}>
      <Text style={styles.formatEmoji}>{format === 'pdf' ? '📄' : '📊'}</Text>
      <Text style={styles.formatLabel}>{format === 'pdf' ? 'PDF' : 'Excel'}</Text>
    </View>
  );
}

export function ExportCard({
  item,
  selected,
  onSelectToggle,
  onSharePdf,
  onShareXlsx,
  onDelete,
  sharing
}: Props) {
  const hasPdf = Boolean(item.pdfPath || item.pdfFilename);
  const hasXlsx = Boolean(item.xlsxPath || item.xlsxFilename);

  return (
    <Card style={selected ? { ...styles.card, ...styles.selected } : styles.card}>
      <View style={styles.header}>
        {onSelectToggle ? <SelectionCheckbox selected={selected} onToggle={onSelectToggle} /> : null}
        <View style={styles.icons}>
          {hasPdf ? <FormatIcon format="pdf" /> : null}
          {hasXlsx ? <FormatIcon format="xlsx" /> : null}
        </View>
        <View style={styles.info}>
          <Text style={styles.title} numberOfLines={2}>
            {item.protocolTitle || item.projectName}
          </Text>
          <Text style={styles.meta}>
            {item.projectName} · {item.protocolDate}
          </Text>
        </View>
      </View>
      {!onSelectToggle ? (
        <View style={styles.actions}>
          {hasPdf && onSharePdf ? (
            <Pressable
              style={({ pressed }) => [styles.actionBtn, pressed ? styles.actionPressed : null]}
              disabled={sharing}
              onPress={() => {
                void hapticSelection();
                onSharePdf();
              }}
            >
              <Text style={styles.actionLabel}>{sharing ? 'Teilen…' : 'PDF teilen'}</Text>
            </Pressable>
          ) : null}
          {hasXlsx && onShareXlsx ? (
            <Pressable
              style={({ pressed }) => [styles.actionBtn, pressed ? styles.actionPressed : null]}
              disabled={sharing}
              onPress={() => {
                void hapticSelection();
                onShareXlsx();
              }}
            >
              <Text style={styles.actionLabel}>{sharing ? 'Teilen…' : 'Excel teilen'}</Text>
            </Pressable>
          ) : null}
          {onDelete ? (
            <Pressable
              style={({ pressed }) => [styles.actionBtn, styles.deleteBtn, pressed ? styles.actionPressed : null]}
              onPress={() => {
                void hapticSelection();
                onDelete();
              }}
            >
              <Text style={[styles.actionLabel, styles.deleteLabel]}>Löschen</Text>
            </Pressable>
          ) : null}
        </View>
      ) : null}
    </Card>
  );
}

const styles = StyleSheet.create({
  card: {
    marginBottom: spacing.sm,
    ...shadows.card
  },
  selected: {
    borderColor: colors.accent,
    borderWidth: 2
  },
  header: {
    flexDirection: 'row',
    gap: spacing.sm,
    alignItems: 'flex-start'
  },
  icons: {
    flexDirection: 'row',
    gap: spacing.xs
  },
  formatIcon: {
    width: 52,
    height: 52,
    borderRadius: spacing.iconRadius,
    alignItems: 'center',
    justifyContent: 'center',
    gap: 2
  },
  pdfIcon: {
    backgroundColor: 'rgba(196, 75, 50, 0.12)'
  },
  xlsxIcon: {
    backgroundColor: 'rgba(47, 107, 69, 0.12)'
  },
  formatEmoji: {
    fontSize: 18
  },
  formatLabel: {
    ...typography.caption,
    fontSize: 10,
    color: colors.muted
  },
  info: {
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
  actions: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs,
    marginTop: spacing.md,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border
  },
  actionBtn: {
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.md,
    borderRadius: spacing.buttonRadius,
    backgroundColor: colors.bg,
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: colors.border
  },
  deleteBtn: {
    backgroundColor: 'transparent',
    borderColor: 'transparent'
  },
  actionPressed: {
    opacity: 0.85
  },
  actionLabel: {
    ...typography.label,
    color: colors.ink
  },
  deleteLabel: {
    color: colors.danger
  }
});
