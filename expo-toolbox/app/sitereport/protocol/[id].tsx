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
  type SiteReportEntry,
  type SiteReportProtocol
} from '../../../src/native/sitereport/db/database';

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
      fields: {
        Kilometer: '',
        Beschreibung: '',
        Status: 'offen'
      },
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
          <TextField
            label="Kilometer"
            value={String(entry.fields.Kilometer ?? '')}
            onChangeText={(value) => {
              const entries = protocol.entries.map((row) =>
                row.id === entry.id ? { ...row, fields: { ...row.fields, Kilometer: value } } : row
              );
              void save({ ...protocol, entries });
            }}
          />
          <TextField
            label="Beschreibung"
            value={String(entry.fields.Beschreibung ?? '')}
            onChangeText={(value) => {
              const entries = protocol.entries.map((row) =>
                row.id === entry.id ? { ...row, fields: { ...row.fields, Beschreibung: value } } : row
              );
              void save({ ...protocol, entries });
            }}
            multiline
          />
          <ListItem
            title="Status"
            subtitle={String(entry.fields.Status ?? '')}
            onPress={() => {
              const nextStatus = entry.fields.Status === 'offen' ? 'erledigt' : 'offen';
              const entries = protocol.entries.map((row) =>
                row.id === entry.id ? { ...row, fields: { ...row.fields, Status: nextStatus } } : row
              );
              void save({ ...protocol, entries });
            }}
          />
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
