import { useCallback, useEffect, useState } from 'react';
import { Alert, Image, StyleSheet, Text, View } from 'react-native';
import * as ImagePicker from 'expo-image-picker';
import { useLocalSearchParams, useRouter } from 'expo-router';

import { PrimaryButton, Screen, TextField } from '../../../../src/components/mobile';
import { useToast } from '../../../../src/contexts/ToastContext';
import { colors, spacing, typography } from '../../../../src/constants/theme';
import { clearLogo, loadLogo, saveLogo, updateProtocol } from '../../../../src/native/sitereport/db/database';
import { getProtocolOrThrow } from '../../../../src/native/sitereport/services/protocolService';

export default function EditProtocolScreen() {
  const router = useRouter();
  const { id } = useLocalSearchParams<{ id: string }>();
  const { showToast } = useToast();
  const [saving, setSaving] = useState(false);
  const [protocolTitle, setProtocolTitle] = useState('');
  const [projectName, setProjectName] = useState('');
  const [description, setDescription] = useState('');
  const [attendees, setAttendees] = useState('');
  const [protocolDate, setProtocolDate] = useState('');
  const [logoDataUrl, setLogoDataUrl] = useState('');
  const [protocolId, setProtocolId] = useState('');

  const load = useCallback(async () => {
    if (!id) return;
    const protocol = await getProtocolOrThrow(id);
    setProtocolId(protocol.id);
    setProtocolTitle(protocol.protocolTitle);
    setProjectName(protocol.projectName);
    setDescription(protocol.protocolDescription);
    setAttendees(protocol.attendees);
    setProtocolDate(protocol.protocolDate);
    setLogoDataUrl(await loadLogo());
  }, [id]);

  useEffect(() => {
    void load();
  }, [load]);

  const pickLogo = async () => {
    const permission = await ImagePicker.requestMediaLibraryPermissionsAsync();
    if (!permission.granted) {
      Alert.alert('Fotos', 'Zugriff auf die Fotobibliothek ist erforderlich.');
      return;
    }
    const result = await ImagePicker.launchImageLibraryAsync({
      mediaTypes: ['images'],
      quality: 0.85,
      base64: true
    });
    if (result.canceled || !result.assets[0]?.base64) return;
    const mime = result.assets[0].mimeType?.includes('png') ? 'image/png' : 'image/jpeg';
    const dataUrl = `data:${mime};base64,${result.assets[0].base64}`;
    await saveLogo(dataUrl);
    setLogoDataUrl(dataUrl);
  };

  const save = async () => {
    if (!protocolId || !protocolTitle.trim()) {
      Alert.alert('Protokoll', 'Bitte einen Protokollnamen eingeben.');
      return;
    }
    setSaving(true);
    try {
      const protocol = await getProtocolOrThrow(protocolId);
      await updateProtocol({
        ...protocol,
        protocolTitle: protocolTitle.trim(),
        projectName: projectName.trim(),
        protocolDescription: description.trim(),
        attendees: attendees.trim()
      });
      showToast('Gespeichert');
      router.back();
    } catch (err) {
      Alert.alert('Fehler', err instanceof Error ? err.message : 'Speichern fehlgeschlagen.');
    } finally {
      setSaving(false);
    }
  };

  return (
    <Screen
      title="Vorgang bearbeiten"
      subtitle="Stammdaten"
      showBack
      footer={<PrimaryButton label={saving ? 'Speichern…' : 'Änderungen speichern'} disabled={saving} onPress={() => void save()} />}
    >
      <TextField label="Protokollname *" value={protocolTitle} onChangeText={setProtocolTitle} />
      <TextField label="Projektname" value={projectName} onChangeText={setProjectName} />
      <TextField label="Beschreibung" value={description} onChangeText={setDescription} multiline />
      <TextField label="Teilnehmer" value={attendees} onChangeText={setAttendees} />
      <View>
        <Text style={styles.label}>Datum</Text>
        <Text style={styles.date}>{protocolDate}</Text>
      </View>

      <Text style={styles.section}>Firmenlogo</Text>
      {logoDataUrl ? (
        <>
          <Image source={{ uri: logoDataUrl }} style={styles.logo} resizeMode="contain" />
          <View style={styles.row}>
            <PrimaryButton label="Logo ändern" variant="secondary" onPress={() => void pickLogo()} />
            <PrimaryButton
              label="Entfernen"
              variant="ghost"
              onPress={() => {
                void clearLogo().then(() => setLogoDataUrl(''));
              }}
            />
          </View>
        </>
      ) : (
        <PrimaryButton label="Logo auswählen" variant="secondary" onPress={() => void pickLogo()} />
      )}
    </Screen>
  );
}

const styles = StyleSheet.create({
  label: { ...typography.label, color: colors.muted, marginBottom: 4 },
  date: { ...typography.bodyStrong, color: colors.ink },
  section: { ...typography.bodyStrong, color: colors.ink, marginTop: spacing.md },
  logo: { width: '100%', height: 96, borderRadius: spacing.inputRadius, backgroundColor: colors.panel },
  row: { flexDirection: 'row', flexWrap: 'wrap', gap: spacing.xs }
});
