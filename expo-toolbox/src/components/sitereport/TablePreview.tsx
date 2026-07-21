import { ScrollView, StyleSheet, Text, View } from 'react-native';

import { colors, spacing, typography } from '../../constants/theme';
import type { SiteReportColumn } from '../../native/sitereport/db/database';
import { Card } from '../mobile';

type Props = {
  columns: SiteReportColumn[];
};

export function TablePreview({ columns }: Props) {
  return (
    <Card style={styles.card}>
      <Text style={styles.label}>Tabellenvorschau</Text>
      <ScrollView horizontal showsHorizontalScrollIndicator={false}>
        <View>
          <View style={styles.row}>
            {columns.map((col) => (
              <View key={col.id} style={styles.cellHeader}>
                <Text style={styles.headerText} numberOfLines={1}>
                  {col.name}
                </Text>
              </View>
            ))}
          </View>
          <View style={styles.row}>
            {columns.map((col) => (
              <View key={`${col.id}-sample`} style={styles.cell}>
                <Text style={styles.cellText} numberOfLines={1}>
                  {col.isPhoto ? 'Foto' : col.type === 'number' ? '123' : 'Beispiel'}
                </Text>
              </View>
            ))}
          </View>
        </View>
      </ScrollView>
    </Card>
  );
}

const styles = StyleSheet.create({
  card: {
    gap: spacing.sm
  },
  label: {
    ...typography.label,
    color: colors.muted
  },
  row: {
    flexDirection: 'row'
  },
  cellHeader: {
    minWidth: 100,
    padding: spacing.sm,
    backgroundColor: colors.accent2,
    borderWidth: 1,
    borderColor: colors.border
  },
  headerText: {
    ...typography.label,
    color: colors.white
  },
  cell: {
    minWidth: 100,
    padding: spacing.sm,
    backgroundColor: colors.panelElevated,
    borderWidth: 1,
    borderColor: colors.border
  },
  cellText: {
    ...typography.caption,
    color: colors.ink
  }
});
