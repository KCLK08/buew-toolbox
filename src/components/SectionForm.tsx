import { Pressable, StyleSheet, Text, View } from 'react-native';

import { FieldInput } from './FieldInput';
import { PhotoDocEditor } from './PhotoDocEditor';
import { colors } from '@/theme/colors';
import type { PhotoDoc, RunSection } from '@/types';
import { inputKeyForCell, inputKeyForField } from '@/lib/setup-model';
import { getVisibleRowCount, tableRowCountKey } from '@/lib/run-utils';
import { normalizeClockTime } from '@/lib/time-format';

interface SectionFormProps {
  section: RunSection;
  values: Record<string, string | boolean>;
  photoDoc: PhotoDoc;
  onValueChange: (key: string, value: string | boolean) => void;
  onPhotoDocChange: (photoDoc: PhotoDoc) => void;
  onWeatherSync?: () => void;
  weatherSyncBusy?: boolean;
}

export function SectionForm({
  section,
  values,
  photoDoc,
  onValueChange,
  onPhotoDocChange,
  onWeatherSync,
  weatherSyncBusy = false,
}: SectionFormProps) {
  if (section.kind === 'photo-doc') {
    return <PhotoDocEditor photoDoc={photoDoc} onChange={onPhotoDocChange} />;
  }

  if (section.kind === 'single') {
    return (
      <View>
        {section.sectionId === 'single:weather' && onWeatherSync ? (
          <Pressable style={styles.weatherButton} onPress={onWeatherSync} disabled={weatherSyncBusy}>
            <Text style={styles.weatherButtonText}>{weatherSyncBusy ? 'Wetter wird geladen…' : 'Wetter automatisch laden'}</Text>
          </Pressable>
        ) : null}
        {(section.fields || [])
          .filter((field) => !field.skipped)
          .map((field) => {
            const key = inputKeyForField(field);
            return (
              <FieldInput
                key={key}
                field={field}
                value={values[key] as string | boolean | undefined}
                onChange={(value) => onValueChange(key, value)}
              />
            );
          })}
      </View>
    );
  }

  if (section.kind === 'table') {
    const tableId = section.tableId || '';
    const rows = section.rows || [];
    const visibleCount = getVisibleRowCount(tableId, values, rows.length);
    const visibleRows = rows.slice(0, visibleCount);

    return (
      <View>
        {visibleRows.map((row) => (
          <View key={row.rowId} style={styles.tableRow}>
            <Text style={styles.rowTitle}>Zeile {row.index}</Text>
            {(row.cells || [])
              .filter((cell) => !cell.skipped)
              .map((cell) => {
                const key = inputKeyForCell(cell);
                const column = (section.columns || []).find((c) => c.columnId === cell.columnId);
                const label = column?.label || cell.label;
                const isTime = tableId === 'table_main_personal' && ['c2', 'c3'].includes(cell.columnId);
                return (
                  <FieldInput
                    key={key}
                    field={{ ...cell, label }}
                    value={values[key] as string | boolean | undefined}
                    onChange={(value) => onValueChange(key, isTime ? normalizeClockTime(String(value)) : value)}
                    compact
                  />
                );
              })}
          </View>
        ))}

        <View style={styles.rowActions}>
          {visibleCount < rows.length ? (
            <Pressable
              style={styles.rowButton}
              onPress={() => onValueChange(tableRowCountKey(tableId), String(visibleCount + 1))}
            >
              <Text style={styles.rowButtonText}>+ Zeile hinzufügen</Text>
            </Pressable>
          ) : null}
          {visibleCount > 1 ? (
            <Pressable
              style={[styles.rowButton, styles.rowButtonSecondary]}
              onPress={() => onValueChange(tableRowCountKey(tableId), String(visibleCount - 1))}
            >
              <Text style={[styles.rowButtonText, styles.rowButtonTextSecondary]}>Zeile entfernen</Text>
            </Pressable>
          ) : null}
        </View>
      </View>
    );
  }

  return null;
}

const styles = StyleSheet.create({
  weatherButton: {
    backgroundColor: '#e8f4f2',
    borderColor: colors.primary,
    borderRadius: 10,
    borderWidth: 1,
    marginBottom: 16,
    paddingVertical: 12,
  },
  weatherButtonText: {
    color: colors.primary,
    fontWeight: '600',
    textAlign: 'center',
  },
  tableRow: {
    backgroundColor: colors.surface,
    borderColor: colors.border,
    borderRadius: 12,
    borderWidth: 1,
    marginBottom: 12,
    padding: 12,
  },
  rowTitle: {
    color: colors.accent,
    fontSize: 13,
    fontWeight: '700',
    marginBottom: 8,
  },
  rowActions: {
    flexDirection: 'row',
    gap: 10,
  },
  rowButton: {
    backgroundColor: colors.primary,
    borderRadius: 10,
    flex: 1,
    paddingVertical: 10,
  },
  rowButtonSecondary: {
    backgroundColor: colors.surface,
    borderColor: colors.border,
    borderWidth: 1,
  },
  rowButtonText: {
    color: '#fff',
    fontWeight: '600',
    textAlign: 'center',
  },
  rowButtonTextSecondary: {
    color: colors.textMuted,
  },
});
