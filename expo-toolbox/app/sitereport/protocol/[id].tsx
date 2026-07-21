import { useCallback, useEffect, useState } from 'react';
import { Alert, Image, StyleSheet, Text, View } from 'react-native';
import * as ImagePicker from 'expo-image-picker';
import { useLocalSearchParams, useRouter } from 'expo-router';

import { ListItem, PrimaryButton, Screen, TextField } from '../../../src/components/mobile';
import { colors, typography } from '../../../src/constants/theme';
import {
  exportProtocolPdf,
  exportProtocolXlsx
} from '../../../src/native/sitereport/services/exportService';
import {
  getProtocol,
  updateProtocol,
  type SiteReportColumn,
  type SiteReportEntry,
  type SiteReportProtocol
} from '../../../src/native/sitereport/db/database';

function emptyFieldsFromColumns(columns: SiteReportColumn[]): Record<string, string | number> {
  const fields: Record<string, string | number> = {};
  for (const col of columns) {
    if (!col.isPhoto) {
      fields[col.name] = col.name.toLowerCase() === 'status' ? 'offen' : '';
    }
  }
  return fields;
}

export default function SiteReportProtocolScreen() {
  const router = useRouter();
  const { id } = useLocalSearchParams<{ id: string }>();
  const [protocol, setProtocol] = useState<SiteReportProtocol | null>(null);
  const [exporting, setExporting] = useState<'pdf' | 'xlsx' | null>(null);

  const load = useCallback(async () => {
    if (!id) return;
    setProtocol(await getProtocol(id));
  }, [id]);

  useEffect(() => {
    void load();
  }, [load]);

  const save = async (next: SiteReportProtocol) => {
    await updateProtocol(next);
    setProtocol(next);
  };

  const addEntry = async () => {
    if (!protocol) return;
    const permission = await ImagePicker.requestCameraPermissionsAsync();
    if (!permission.granted) {
      Alert.alert('Kamera', 'Kamerazugriff ist erforderlich.');
      return;
    }
    const result = await ImagePicker.launchCameraAsync({ quality: 0.75 });
    if (result.canceled || !result.assets[0]?.uri) return;

    const entry: SiteReportEntry = {
      id: `entry_${Date.now()}`,
      createdAt: new Date().toISOString(),
      fields: emptyFieldsFromColumns(protocol.columns),
      photoPath: result.assets[0].uri
    };
    await save({ ...protocol, entries: [entry, ...protocol.entries] });
  };

  const runExport = async (format: 'pdf' | 'xlsx') => {
    if (!protocol) return;
    setExporting(format);
    try {
      if (format === 'pdf') {
        await exportProtocolPdf(protocol);
      } else {
        await exportProtocolXlsx(protocol);
      }
      Alert.alert('Export', `${format.toUpperCase()} wurde erstellt und kann geteilt werden.`);
    } catch (err) {
      Alert.alert('Export', err instanceof Error ? err.message : 'Export fehlgeschlagen.');
    } finally {
      setExporting(null);
    }
  };

  const dataColumns = protocol?.columns.filter((col) => !col.isPhoto) ?? [];

  if (!protocol) {
    return (
      <Screen title="SiteReport" showBack>
        <Text style={styles.muted}>Protokoll wird geladen…</Text>
      </Screen>
    );
  }

  return (
    <Screen
      title={protocol.protocolTitle}
      subtitle={protocol.projectName}
      showBack
      footer={
        <View style={styles.exportRow}>
          <PrimaryButton
            label={exporting === 'pdf' ? 'PDF…' : 'PDF exportieren'}
            disabled={Boolean(exporting)}
            onPress={() => void runExport('pdf')}
          />
          <PrimaryButton
            label={exporting === 'xlsx' ? 'XLSX…' : 'XLSX exportieren'}
            variant="secondary"
            disabled={Boolean(exporting)}
            onPress={() => void runExport('xlsx')}
          />
        </View>
      }
    >
      <TextField
        label="Beschreibung"
        value={protocol.protocolDescription}
        onChangeText={(protocolDescription) => void save({ ...protocol, protocolDescription })}
        multiline
      />
      <TextField
        label="Teilnehmer"
        value={protocol.attendees}
        onChangeText={(attendees) => void save({ ...protocol, attendees })}
      />

      <PrimaryButton label="+ Eintrag mit Foto" onPress={() => void addEntry()} />

      <Text style={styles.section}>Einträge ({protocol.entries.length})</Text>
      {protocol.entries.map((entry) => (
        <View key={entry.id} style={styles.entryCard}>
          {entry.photoPath ? <Image source={{ uri: entry.photoPath }} style={styles.photo} /> : null}
          {dataColumns.map((col) => {
            const value = String(entry.fields[col.name] ?? '');
            if (col.name.toLowerCase() === 'status') {
              return (
                <ListItem
                  key={col.id}
                  title={col.name}
                  subtitle={value || 'offen'}
                  onPress={() => {
                    const nextStatus = value === 'offen' ? 'erledigt' : 'offen';
                    const entries = protocol.entries.map((row) =>
                      row.id === entry.id
                        ? { ...row, fields: { ...row.fields, [col.name]: nextStatus } }
                        : row
                    );
                    void save({ ...protocol, entries });
                  }}
                />
              );
            }
            return (
              <TextField
                key={col.id}
                label={col.name}
                value={value}
                keyboardType={col.type === 'number' ? 'decimal-pad' : 'default'}
                onChangeText={(nextValue) => {
                  const entries = protocol.entries.map((row) =>
                    row.id === entry.id ? { ...row, fields: { ...row.fields, [col.name]: nextValue } } : row
                  );
                  void save({ ...protocol, entries });
                }}
                multiline={col.type === 'text' && col.name.length > 12}
              />
            );
          })}
        </View>
      ))}
    </Screen>
  );
}

const styles = StyleSheet.create({
  muted: { ...typography.body, color: colors.muted },
  section: { ...typography.bodyStrong, color: colors.ink, marginTop: 8 },
  entryCard: { gap: 8, marginBottom: 12 },
  photo: { width: '100%', height: 180, borderRadius: 12, backgroundColor: colors.border },
  exportRow: { flexDirection: 'row', gap: 8 }
});
