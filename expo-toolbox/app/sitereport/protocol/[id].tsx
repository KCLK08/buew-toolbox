import { useCallback, useEffect, useState } from 'react';
import { Alert, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter } from 'expo-router';

import { BottomSheet, ConfirmModal, EntryCard } from '../../../src/components/sitereport';
import { PrimaryButton, Screen, StatCard } from '../../../src/components/mobile';
import { useToast } from '../../../src/contexts/ToastContext';
import { colors, spacing, typography } from '../../../src/constants/theme';
import { updateProtocol } from '../../../src/native/sitereport/db/database';
import {
  closeProtocolWithExport,
  exportProtocolPdf,
  exportProtocolXlsx,
  type CloseExportMode
} from '../../../src/native/sitereport/services/exportService';
import {
  getProtocolOrThrow,
  protocolStats,
  removeProtocolEntry
} from '../../../src/native/sitereport/services/protocolService';

export default function SiteReportProtocolScreen() {
  const router = useRouter();
  const { id } = useLocalSearchParams<{ id: string }>();
  const { showToast } = useToast();
  const [protocol, setProtocol] = useState<Awaited<ReturnType<typeof getProtocolOrThrow>> | null>(null);
  const [exporting, setExporting] = useState(false);
  const [closeVisible, setCloseVisible] = useState(false);
  const [exportSheetVisible, setExportSheetVisible] = useState(false);

  const load = useCallback(async () => {
    if (!id) return;
    try {
      setProtocol(await getProtocolOrThrow(id));
    } catch {
      Alert.alert('Fehler', 'Protokoll nicht gefunden.');
      router.back();
    }
  }, [id, router]);

  useEffect(() => {
    void load();
  }, [load]);

  const stats = protocol ? protocolStats(protocol) : null;

  const runExport = async (format: 'pdf' | 'xlsx') => {
    if (!protocol) return;
    setExporting(true);
    setExportSheetVisible(false);
    try {
      if (format === 'pdf') {
        await exportProtocolPdf(protocol);
      } else {
        await exportProtocolXlsx(protocol);
      }
      showToast(`${format.toUpperCase()} exportiert`);
    } catch (err) {
      Alert.alert('Export', err instanceof Error ? err.message : 'Export fehlgeschlagen.');
    } finally {
      setExporting(false);
    }
  };

  const handleClose = async (mode: CloseExportMode) => {
    if (!protocol) return;
    setCloseVisible(false);
    setExporting(true);
    try {
      await updateProtocol(protocol);
      if (mode !== 'save') {
        await closeProtocolWithExport(protocol, mode);
      }
      showToast(mode === 'save' ? 'Gespeichert' : 'Protokoll abgeschlossen');
      router.replace('/sitereport/protocols');
    } catch (err) {
      Alert.alert('Abschluss', err instanceof Error ? err.message : 'Abschluss fehlgeschlagen.');
    } finally {
      setExporting(false);
    }
  };

  const deleteEntry = (entryId: string) => {
    if (!protocol) return;
    Alert.alert('Eintrag löschen', 'Diesen Eintrag wirklich entfernen?', [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Löschen',
        style: 'destructive',
        onPress: () => {
          void removeProtocolEntry(protocol, entryId).then((next) => {
            setProtocol(next);
            showToast('Eintrag gelöscht');
          });
        }
      }
    ]);
  };

  if (!protocol || !stats) {
    return (
      <Screen title="Protokoll" showBack>
        <Text style={styles.muted}>Protokoll wird geladen…</Text>
      </Screen>
    );
  }

  return (
    <Screen
      title={protocol.protocolTitle}
      subtitle={protocol.projectName}
      showBack
      scroll
      onRefresh={load}
      refreshing={exporting}
      footer={
        <View style={styles.footer}>
          <PrimaryButton
            label="+ Neuer Eintrag"
            onPress={() => router.push(`/sitereport/protocol/${protocol.id}/wizard`)}
          />
        </View>
      }
    >
      <View style={styles.statsRow}>
        <StatCard title="Datum" value={protocol.protocolDate} icon="📅" />
        <StatCard title="Einträge" value={String(stats.entryCount)} icon="✏️" />
      </View>
      <View style={styles.statsRow}>
        <StatCard title="Fotos" value={String(stats.photoCount)} icon="📷" />
        <StatCard title="Offen" value={String(stats.openCount)} icon="🟠" tone="warning" />
      </View>

      <View style={styles.actions}>
        <PrimaryButton
          label="Stammdaten bearbeiten"
          variant="secondary"
          onPress={() => router.push(`/sitereport/protocol/${protocol.id}/edit`)}
        />
        <PrimaryButton
          label={exporting ? 'Export…' : 'Export'}
          variant="secondary"
          disabled={exporting}
          onPress={() => setExportSheetVisible(true)}
        />
        <PrimaryButton
          label="Protokoll abschließen"
          variant="ghost"
          disabled={exporting}
          onPress={() => setCloseVisible(true)}
        />
      </View>

      <Text style={styles.section}>Einträge ({protocol.entries.length})</Text>
      {protocol.entries.length === 0 ? (
        <Text style={styles.muted}>Noch keine Einträge. Tippe auf „+ Neuer Eintrag".</Text>
      ) : (
        protocol.entries.map((entry) => (
          <EntryCard
            key={entry.id}
            entry={entry}
            columns={protocol.columns}
            onEdit={() => router.push(`/sitereport/protocol/${protocol.id}/wizard?entryId=${entry.id}`)}
            onDelete={() => deleteEntry(entry.id)}
          />
        ))
      )}

      <ConfirmModal
        visible={closeVisible}
        title="Protokoll abschließen"
        message="Wie möchtest du fortfahren?"
        onClose={() => setCloseVisible(false)}
        actions={[
          { label: 'Nur speichern', onPress: () => void handleClose('save'), variant: 'secondary' },
          { label: 'PDF erstellen', onPress: () => void handleClose('pdf'), variant: 'primary' },
          { label: 'Excel erstellen', onPress: () => void handleClose('xlsx'), variant: 'primary' },
          { label: 'PDF + Excel', onPress: () => void handleClose('both'), variant: 'primary' },
          { label: 'Abbrechen', onPress: () => setCloseVisible(false), variant: 'ghost' }
        ]}
      />

      <BottomSheet visible={exportSheetVisible} title="Export" onClose={() => setExportSheetVisible(false)}>
        <PrimaryButton label="PDF exportieren" onPress={() => void runExport('pdf')} disabled={exporting} />
        <PrimaryButton
          label="Excel exportieren"
          variant="secondary"
          onPress={() => void runExport('xlsx')}
          disabled={exporting}
        />
      </BottomSheet>
    </Screen>
  );
}

const styles = StyleSheet.create({
  muted: { ...typography.body, color: colors.muted },
  section: { ...typography.bodyStrong, color: colors.ink, marginTop: spacing.sm, marginBottom: spacing.xs },
  statsRow: { flexDirection: 'row', gap: spacing.sm, marginBottom: spacing.sm },
  actions: { gap: spacing.xs, marginBottom: spacing.md },
  footer: { gap: spacing.xs }
});
