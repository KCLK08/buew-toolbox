import { useCallback, useEffect, useState } from 'react';
import { Alert, Image, StyleSheet, Text, View } from 'react-native';
import * as ImagePicker from 'expo-image-picker';
import { useRouter } from 'expo-router';

import { TablePreview, WizardFooter, WizardStep } from '../../src/components/sitereport';
import { ListItem, Screen, TextField } from '../../src/components/mobile';
import { useToast } from '../../src/contexts/ToastContext';
import { colors, spacing, typography } from '../../src/constants/theme';
import { hapticLight, hapticSuccess } from '../../src/lib/haptics';
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
    void hapticLight();
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

  const goNext = () => {
    void hapticLight();
    setStep((s) => s + 1);
  };

  const goBack = () => {
    void hapticLight();
    setStep((s) => s - 1);
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
      void hapticSuccess();
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
      showBack
      footer={
        <WizardFooter
          showBack={step > 1}
          onBack={goBack}
          primaryLabel={step < 3 ? 'Weiter' : saving ? 'Erstellen…' : 'Protokoll erstellen'}
          onPrimary={step < 3 ? goNext : () => void create()}
          primaryDisabled={step < 3 ? !canNext() : saving || !canNext()}
          primaryLoading={saving}
        />
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
              autoFocus
            />
            <TextField label="Projektname" value={projectName} onChangeText={setProjectName} placeholder="Baustelle Nord" />
            <TextField label="Beschreibung" value={description} onChangeText={setDescription} multiline placeholder="Optional" />
            <TextField label="Teilnehmer" value={attendees} onChangeText={setAttendees} placeholder="Optional" />
            <View style={styles.dateCard}>
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
                <View style={styles.row}>
                  <Text style={styles.logoAction} onPress={() => void pickLogo()}>
                    Logo ändern
                  </Text>
                  <Text style={styles.logoRemove} onPress={() => void removeLogo()}>
                    Entfernen
                  </Text>
                </View>
              </>
            ) : (
              <View style={styles.logoPlaceholder}>
                <Text style={styles.logoIcon}>🏢</Text>
                <Text style={styles.logoHint} onPress={() => void pickLogo()}>
                  Logo auswählen
                </Text>
              </View>
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
                meta={selectedTemplateId === tpl.id ? '✓ Ausgewählt' : undefined}
                onPress={() => void selectTemplate(tpl.id)}
              />
            ))}
            <View style={styles.row}>
              <Text style={styles.linkAction} onPress={() => router.push('/sitereport/format-builder?mode=new')}>
                + Neues Format
              </Text>
              {selectedTemplateId ? (
                <Text
                  style={styles.linkAction}
                  onPress={() =>
                    router.push(`/sitereport/format-builder?mode=edit&templateId=${selectedTemplateId}`)
                  }
                >
                  Bearbeiten
                </Text>
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
  gap: { gap: spacing.md },
  row: { flexDirection: 'row', flexWrap: 'wrap', gap: spacing.md },
  label: { ...typography.label, color: colors.muted, marginBottom: 4 },
  dateCard: {
    backgroundColor: colors.panelElevated,
    borderRadius: spacing.inputRadius,
    borderWidth: 1,
    borderColor: colors.border,
    padding: spacing.md
  },
  dateValue: { ...typography.bodyStrong, color: colors.ink },
  hint: { ...typography.body, color: colors.muted },
  logo: {
    width: '100%',
    height: 160,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panel
  },
  logoPlaceholder: {
    height: 160,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.panel,
    borderWidth: 2,
    borderColor: colors.border,
    borderStyle: 'dashed',
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm
  },
  logoIcon: { fontSize: 40 },
  logoHint: { ...typography.bodyStrong, color: colors.accent },
  logoAction: { ...typography.bodyStrong, color: colors.accent },
  logoRemove: { ...typography.bodyStrong, color: colors.danger },
  linkAction: { ...typography.bodyStrong, color: colors.accent }
});
