import { useCallback, useEffect, useState } from 'react';
import { Alert, Image, StyleSheet, Text, View } from 'react-native';
import * as ImagePicker from 'expo-image-picker';
import { useRouter } from 'expo-router';

import { EmptyState, Fab, ListItem, PrimaryButton, Screen, TextField } from '../../../src/components/mobile';
import { colors, typography } from '../../../src/constants/theme';
import {
  clearLogo,
  createProtocol,
  getActiveColumns,
  initSiteReportDatabase,
  listExports,
  listProtocols,
  listTemplates,
  loadLogo,
  loadSettings,
  saveLogo,
  saveSettings,
  todayDe,
  type SiteReportColumn,
  type SiteReportExport,
  type SiteReportTemplate
} from '../../../src/native/sitereport/db/database';
import { deleteCachedExport, shareCachedExport } from '../../../src/native/sitereport/services/exportService';

export default function SiteReportHomeScreen() {
  const router = useRouter();
  const [loading, setLoading] = useState(true);
  const [protocols, setProtocols] = useState<Awaited<ReturnType<typeof listProtocols>>>([]);
  const [templates, setTemplates] = useState<SiteReportTemplate[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState('');
  const [columns, setColumns] = useState<SiteReportColumn[]>([]);
  const [logoDataUrl, setLogoDataUrl] = useState('');
  const [title, setTitle] = useState('');
  const [project, setProject] = useState('');
  const [exportsList, setExportsList] = useState<SiteReportExport[]>([]);
  const [sharingExport, setSharingExport] = useState<string | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    await initSiteReportDatabase();
    const [nextProtocols, nextTemplates, settings, logo, nextExports] = await Promise.all([
      listProtocols(),
      listTemplates(),
      loadSettings(),
      loadLogo(),
      listExports()
    ]);
    setProtocols(nextProtocols);
    setTemplates(nextTemplates);
    setExportsList(nextExports);
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
    setLoading(false);
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
  };

  const removeLogo = async () => {
    await clearLogo();
    setLogoDataUrl('');
  };

  const startProtocol = async () => {
    if (!selectedTemplateId) {
      Alert.alert('Format', 'Bitte zuerst ein Tabellenformat auswählen.');
      return;
    }
    const protocol = await createProtocol({
      protocolTitle: title.trim() || 'Neues Protokoll',
      projectName: project.trim() || 'Projekt',
      protocolDate: todayDe(),
      columns
    });
    setTitle('');
    setProject('');
    router.push(`/sitereport/protocol/${protocol.id}`);
  };

  const shareExport = async (exportId: string, format: 'pdf' | 'xlsx') => {
    setSharingExport(`${exportId}:${format}`);
    try {
      await shareCachedExport(exportId, format);
    } catch (err) {
      Alert.alert('Export', err instanceof Error ? err.message : 'Teilen fehlgeschlagen.');
    } finally {
      setSharingExport(null);
    }
  };

  const removeExport = (exportId: string) => {
    Alert.alert('Export löschen', 'Gespeicherten Export wirklich entfernen?', [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Löschen',
        style: 'destructive',
        onPress: () => {
          void deleteCachedExport(exportId).then(load);
        }
      }
    ]);
  };

  return (
    <Screen title="SiteReport" subtitle="Foto-Protokolle mit Export" scroll refreshing={loading} onRefresh={load}>
      <View style={styles.card}>
        <Text style={styles.cardTitle}>Firmenlogo (optional)</Text>
        {logoDataUrl ? (
          <>
            <Image source={{ uri: logoDataUrl }} style={styles.logo} resizeMode="contain" />
            <View style={styles.row}>
              <PrimaryButton label="Logo ändern" variant="secondary" onPress={() => void pickLogo()} />
              <PrimaryButton label="Logo entfernen" variant="ghost" onPress={() => void removeLogo()} />
            </View>
          </>
        ) : (
          <PrimaryButton label="Logo hochladen" variant="secondary" onPress={() => void pickLogo()} />
        )}
      </View>

      <View style={styles.card}>
        <Text style={styles.cardTitle}>Tabellenformat</Text>
        {templates.length === 0 ? (
          <Text style={styles.muted}>Noch kein Format vorhanden.</Text>
        ) : (
          templates.map((tpl) => (
            <ListItem
              key={tpl.id}
              title={tpl.name}
              subtitle={`${tpl.columns.length} Spalten`}
              meta={selectedTemplateId === tpl.id ? 'Aktiv' : undefined}
              onPress={() => void selectTemplate(tpl.id)}
            />
          ))
        )}
        <View style={styles.row}>
          <PrimaryButton
            label="Neues Format"
            variant="secondary"
            onPress={() => router.push('/sitereport/format-builder?mode=new')}
          />
          {selectedTemplateId ? (
            <PrimaryButton
              label="Format bearbeiten"
              variant="ghost"
              onPress={() => router.push(`/sitereport/format-builder?mode=edit&templateId=${selectedTemplateId}`)}
            />
          ) : null}
        </View>
      </View>

      <View style={styles.card}>
        <Text style={styles.cardTitle}>Neues Protokoll</Text>
        <TextField label="Titel" value={title} onChangeText={setTitle} />
        <TextField label="Projekt" value={project} onChangeText={setProject} />
        <PrimaryButton label="Protokoll starten" disabled={!selectedTemplateId} onPress={() => void startProtocol()} />
      </View>

      {exportsList.length > 0 ? (
        <View style={styles.card}>
          <Text style={styles.cardTitle}>Exporte ({exportsList.length})</Text>
          {exportsList.map((item) => (
            <View key={item.id} style={styles.exportCard}>
              <Text style={styles.exportTitle}>{item.protocolTitle || item.projectName}</Text>
              <Text style={styles.muted}>
                {item.projectName} · {item.protocolDate}
              </Text>
              <View style={styles.row}>
                {item.pdfPath || item.pdfFilename ? (
                  <PrimaryButton
                    label={sharingExport === `${item.id}:pdf` ? 'PDF…' : 'PDF teilen'}
                    variant="secondary"
                    disabled={Boolean(sharingExport)}
                    onPress={() => void shareExport(item.id, 'pdf')}
                  />
                ) : null}
                {item.xlsxPath || item.xlsxFilename ? (
                  <PrimaryButton
                    label={sharingExport === `${item.id}:xlsx` ? 'XLSX…' : 'XLSX teilen'}
                    variant="secondary"
                    disabled={Boolean(sharingExport)}
                    onPress={() => void shareExport(item.id, 'xlsx')}
                  />
                ) : null}
                <PrimaryButton label="Löschen" variant="ghost" onPress={() => removeExport(item.id)} />
              </View>
            </View>
          ))}
        </View>
      ) : null}

      <Text style={styles.section}>Protokolle</Text>
      {protocols.length === 0 ? (
        <EmptyState title="Keine Protokolle" description="Erstelle ein neues Foto-Protokoll für die Baustelle." />
      ) : (
        protocols.map((protocol) => (
          <ListItem
            key={protocol.id}
            title={protocol.protocolTitle}
            subtitle={protocol.projectName}
            meta={`${protocol.protocolDate} · ${protocol.entries.length} Einträge`}
            onPress={() => router.push(`/sitereport/protocol/${protocol.id}`)}
          />
        ))
      )}

      <Fab label="+" onPress={() => void startProtocol()} accessibilityLabel="Neues Protokoll" />
    </Screen>
  );
}

const styles = StyleSheet.create({
  card: { gap: 12, marginBottom: 16 },
  cardTitle: { ...typography.bodyStrong, color: colors.ink },
  section: { ...typography.bodyStrong, color: colors.ink, marginBottom: 8 },
  exportCard: { gap: 8, padding: 12, borderRadius: 12, backgroundColor: colors.panel, borderWidth: 1, borderColor: colors.border },
  exportTitle: { ...typography.bodyStrong, color: colors.ink },
  muted: { ...typography.body, color: colors.muted },
  logo: { width: '100%', height: 96, borderRadius: 12, backgroundColor: colors.panel },
  row: { flexDirection: 'row', flexWrap: 'wrap', gap: 8 }
});
