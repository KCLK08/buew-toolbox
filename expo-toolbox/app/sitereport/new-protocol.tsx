import { useCallback, useEffect, useState } from 'react';
import { Alert, Image, ScrollView, StyleSheet, Text, View } from 'react-native';
import * as ImagePicker from 'expo-image-picker';
import { useRouter } from 'expo-router';

import { TablePreview, WizardStep } from '../../src/components/sitereport';
import { ListItem, PrimaryButton, Screen, TextField } from '../../src/components/mobile';
import { useToast } from '../../src/contexts/ToastContext';
import { colors, spacing, typography } from '../../src/constants/theme';
import {
  clearLogo,
  createProtocol,
  getActiveColumns,
  initSiteReportDatabase,
  listTemplates,
  loadLogo,
  loadSettings,
  saveLogo,
  saveSettings,
  todayDe,
  type SiteReportColumn,
  type SiteReportTemplate
} from '../../src/native/sitereport/db/database';

const STEPS = ['Allgemeine Daten', 'Firmenlogo', 'Tabellenformat'] as const;

export default function NewProtocolScreen() {
  const router = useRouter();
  const { showToast } = useToast();
  const [step, setStep] = useState(1);
  const [saving, setSaving] = useState(false);

  const [protocolTitle, setProtocolTitle] = useState('');
  const [projectName, setProjectName] = useState('');
  const [description, setDescription] = useState('');
  const [attendees, setAttendees] = useState('');
  const [protocolDate] = useState(todayDe());

  const [logoDataUrl, setLogoDataUrl] = useState('');
  const [templates, setTemplates] = useState<SiteReportTemplate[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState('');
  const [columns, setColumns] = useState<SiteReportColumn[]>([]);

  const load = useCallback(async () => {
    await initSiteReportDatabase();
    const [nextTemplates, settings, logo] = await Promise.all([
      listTemplates(),
      loadSettings(),
      loadLogo()
    ]);
    setTemplates(nextTemplates);
    setLogoDataUrl(logo);
    const templateId = settings?.selectedTemplateId || nextTemplates[0]?.id || '';
    setSelectedTemplateId(templateId);
    if (settings?.columns?.length) {
      setColumns(settings.columns);
    } else if (templateId) {
      const tpl = nextTemplates.find((row) => row.id === templateId);
      setColumns(tpl?.columns ?? (await getActiveColumns()));
    } else {
      setColumns(await getActiveColumns());
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const selectTemplate = async (templateId: string) => {
    const tpl = templates.find((row) => row.id === templateId);
    if (!tpl) return;
    setSelectedTemplateId(templateId);
    setColumns(tpl.columns);
    await saveSettings({ selectedTemplateId: templateId, columns: tpl.columns });
  };

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
    showToast('Logo gespeichert');
  };

  const removeLogo = async () => {
    await clearLogo();
    setLogoDataUrl('');
    showToast('Logo entfernt');
  };

  const canNext = () => {
    if (step === 1) return protocolTitle.trim().length > 0;
    if (step === 3) return Boolean(selectedTemplateId);
    return true;
  };

  const create = async () => {
    if (!selectedTemplateId) {
      Alert.alert('Format', 'Bitte ein Tabellenformat auswählen.');
      return;
    }
    if (!protocolTitle.trim()) {
      Alert.alert('Protokoll', 'Bitte einen Protokollnamen eingeben.');
      return;
    }
    setSaving(true);
    try {
      const protocol = await createProtocol({
        protocolTitle: protocolTitle.trim(),
        projectName: projectName.trim() || 'Projekt',
        protocolDate,
        protocolDescription: description.trim(),
        attendees: attendees.trim(),
        columns
      });
      showToast('Protokoll erstellt');
      router.replace(`/sitereport/protocol/${protocol.id}`);
    } catch (err) {
      Alert.alert('Fehler', err instanceof Error ? err.message : 'Erstellen fehlgeschlagen.');
    } finally {
      setSaving(false);
    }
  };

  return (
    <Screen
      title="Neues Protokoll"
      subtitle={STEPS[step - 1]}
      showBack
      footer={
        <View style={styles.footer}>
          {step > 1 ? (
            <PrimaryButton label="Zurück" variant="secondary" onPress={() => setStep((s) => s - 1)} />
          ) : null}
          {step < 3 ? (
            <PrimaryButton
              label="Weiter"
              disabled={!canNext()}
              onPress={() => setStep((s) => s + 1)}
            />
          ) : (
            <PrimaryButton
              label={saving ? 'Erstellen…' : 'Protokoll erstellen'}
              disabled={saving || !canNext()}
              onPress={() => void create()}
            />
          )}
        </View>
      }
    >
      <WizardStep step={step} total={3} title={STEPS[step - 1]}>
        {step === 1 ? (
          <View style={styles.gap}>
            <TextField
              label="Protokollname *"
              value={protocolTitle}
              onChangeText={setProtocolTitle}
              placeholder="z. B. Tagesprotokoll"
            />
            <TextField label="Projektname" value={projectName} onChangeText={setProjectName} />
            <TextField label="Beschreibung" value={description} onChangeText={setDescription} multiline />
            <TextField label="Teilnehmer" value={attendees} onChangeText={setAttendees} />
            <View>
              <Text style={styles.label}>Datum</Text>
              <Text style={styles.dateValue}>{protocolDate}</Text>
            </View>
          </View>
        ) : null}

        {step === 2 ? (
          <View style={styles.gap}>
            <Text style={styles.hint}>Optional — wird in PDF und Excel eingebettet.</Text>
            {logoDataUrl ? (
              <>
                <Image source={{ uri: logoDataUrl }} style={styles.logo} resizeMode="contain" />
                <PrimaryButton label="Logo ändern" variant="secondary" onPress={() => void pickLogo()} />
                <PrimaryButton label="Logo entfernen" variant="ghost" onPress={() => void removeLogo()} />
              </>
            ) : (
              <PrimaryButton label="Logo auswählen" onPress={() => void pickLogo()} />
            )}
          </View>
        ) : null}

        {step === 3 ? (
          <View style={styles.gap}>
            {templates.map((tpl) => (
              <ListItem
                key={tpl.id}
                title={tpl.name}
                subtitle={`${tpl.columns.length} Spalten`}
                meta={selectedTemplateId === tpl.id ? 'Aktiv' : undefined}
                onPress={() => void selectTemplate(tpl.id)}
              />
            ))}
            <View style={styles.row}>
              <PrimaryButton
                label="Neues Format"
                variant="secondary"
                onPress={() => router.push('/sitereport/format-builder?mode=new')}
              />
              {selectedTemplateId ? (
                <PrimaryButton
                  label="Bearbeiten"
                  variant="ghost"
                  onPress={() =>
                    router.push(`/sitereport/format-builder?mode=edit&templateId=${selectedTemplateId}`)
                  }
                />
              ) : null}
            </View>
            {columns.length > 0 ? <TablePreview columns={columns} /> : null}
          </View>
        ) : null}
      </WizardStep>
    </Screen>
  );
}

const styles = StyleSheet.create({
  gap: { gap: spacing.sm },
  footer: { gap: spacing.xs },
  row: { flexDirection: 'row', flexWrap: 'wrap', gap: spacing.xs },
  label: { ...typography.label, color: colors.muted, marginBottom: 4 },
  dateValue: { ...typography.bodyStrong, color: colors.ink },
  hint: { ...typography.body, color: colors.muted },
  logo: { width: '100%', height: 120, borderRadius: spacing.inputRadius, backgroundColor: colors.panel }
});
