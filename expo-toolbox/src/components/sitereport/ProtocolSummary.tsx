import { Image, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import type { SiteReportColumn, SiteReportEntry } from '../../native/sitereport/db/database';
import { Card, StatusBadge } from '../mobile';

type Props = {
  entry: SiteReportEntry;
  columns: SiteReportColumn[];
};

export function ProtocolSummary({ entry, columns }: Props) {
  const dataColumns = columns.filter((col) => !col.isPhoto);

  return (
    <Card style={styles.card}>
      <Text style={styles.heading}>Zusammenfassung</Text>
      {entry.photoPath ? (
        <Image source={{ uri: entry.photoPath }} style={styles.photo} resizeMode="cover" />
      ) : null}
      {dataColumns.map((col) => {
        const value = String(entry.fields[col.name] ?? '—');
        const isStatus = col.name.toLowerCase() === 'status';
        return (
          <View key={col.id} style={styles.row}>
            <Text style={styles.label}>{col.name}</Text>
            {isStatus ? (
              <StatusBadge label={value} tone={value === 'erledigt' ? 'success' : value === 'bearbeitung' ? 'info' : 'warning'} />
            ) : (
              <Text style={styles.value}>{value}</Text>
            )}
          </View>
        );
      })}
    </Card>
  );
}

const styles = StyleSheet.create({
  card: {
    gap: spacing.sm
  },
  heading: {
    ...typography.subtitle,
    color: colors.ink
  },
  photo: {
    width: '100%',
    height: 180,
    borderRadius: spacing.inputRadius,
    backgroundColor: colors.border
  },
  row: {
    gap: 4,
    paddingVertical: 4
  },
  label: {
    ...typography.label,
    color: colors.muted
  },
  value: {
    ...typography.body,
    color: colors.ink
  }
});
