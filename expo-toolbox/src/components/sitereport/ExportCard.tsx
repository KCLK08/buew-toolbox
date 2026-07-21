import { StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import type { SiteReportExport } from '../../native/sitereport/db/database';
import { Card, PrimaryButton } from '../mobile';

type Props = {
  item: SiteReportExport;
  selected?: boolean;
  onSelectToggle?: () => void;
  onSharePdf?: () => void;
  onShareXlsx?: () => void;
  onDelete?: () => void;
  sharing?: boolean;
};

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
        {onSelectToggle ? (
          <PrimaryButton
            label={selected ? '☑' : '☐'}
            variant="ghost"
            onPress={onSelectToggle}
          />
        ) : null}
        <View style={styles.info}>
          <Text style={styles.title}>{item.protocolTitle || item.projectName}</Text>
          <Text style={styles.meta}>
            {item.projectName} · {item.protocolDate}
          </Text>
          <View style={styles.badges}>
            {hasPdf ? <Text style={styles.badge}>PDF</Text> : null}
            {hasXlsx ? <Text style={styles.badge}>Excel</Text> : null}
          </View>
        </View>
      </View>
      <View style={styles.actions}>
        {hasPdf && onSharePdf ? (
          <PrimaryButton
            label={sharing ? 'PDF…' : 'PDF teilen'}
            variant="secondary"
            disabled={sharing}
            onPress={onSharePdf}
          />
        ) : null}
        {hasXlsx && onShareXlsx ? (
          <PrimaryButton
            label={sharing ? 'Excel…' : 'Excel teilen'}
            variant="secondary"
            disabled={sharing}
            onPress={onShareXlsx}
          />
        ) : null}
        {onDelete ? <PrimaryButton label="Löschen" variant="ghost" onPress={onDelete} /> : null}
      </View>
    </Card>
  );
}

const styles = StyleSheet.create({
  card: {
    marginBottom: spacing.sm
  },
  selected: {
    borderColor: colors.accent,
    borderWidth: 2
  },
  header: {
    flexDirection: 'row',
    gap: spacing.xs,
    marginBottom: spacing.sm
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
  badges: {
    flexDirection: 'row',
    gap: spacing.xs,
    marginTop: 4
  },
  badge: {
    ...typography.label,
    color: colors.accent,
    backgroundColor: colors.badgeBg,
    paddingHorizontal: spacing.sm,
    paddingVertical: 2,
    borderRadius: 999,
    overflow: 'hidden'
  },
  actions: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs
  }
});
