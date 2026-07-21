import { useCallback, useEffect, useState } from 'react';
import { Alert, Image, StyleSheet, Text, View } from 'react-native';
import * as ImagePicker from 'expo-image-picker';
import { useLocalSearchParams, useRouter } from 'expo-router';

import { ListItem, PrimaryButton, Screen, TextField } from '../../../src/components/mobile';
import { colors, typography } from '../../../src/constants/theme';
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
      footer={<PrimaryButton label="+ Eintrag mit Foto" onPress={() => void addEntry()} />}
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
  photo: { width: '100%', height: 180, borderRadius: 12, backgroundColor: colors.border }
});
