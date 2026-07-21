import { useCallback, useEffect, useState } from 'react';
import { Alert, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter } from 'expo-router';

import {
  ActionChipRow,
  BottomSheet,
  ConfirmModal,
  EntryCard,
  ProtocolHero,
  SectionHeader
} from '../../../src/components/sitereport';
import { EmptyState, PrimaryButton, Screen, StatCard } from '../../../src/components/mobile';
import { useToast } from '../../../src/contexts/ToastContext';
import { spacing } from '../../../src/constants/theme';
import { hapticSuccess } from '../../../src/lib/haptics';
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
      void hapticSuccess();
      showToast(`${format === 'pdf' ? 'PDF' : 'Excel'} wurde erstellt`);
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
      void hapticSuccess();
      showToast(mode === 'save' ? 'Protokoll gespeichert' : 'Protokoll abgeschlossen');
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
        <EmptyState icon="⏳" title="Wird geladen…" description="Protokoll wird vorbereitet." />
      </Screen>
    );
  }

  const statusLabel = stats.openCount > 0 ? `${stats.openCount} offen` : 'Alle erledigt';
  const statusTone = stats.openCount > 0 ? 'warning' : 'success';

  return (
    <Screen
      title="Protokoll"
      showBack
      scroll
      onRefresh={load}
      refreshing={exporting}
      footer={
        <PrimaryButton
          label="+ Neuer Eintrag"
          onPress={() => router.push(`/sitereport/protocol/${protocol.id}/wizard`)}
        />
      }
    >
      <ProtocolHero
        title={protocol.protocolTitle}
        projectName={protocol.projectName}
        date={protocol.protocolDate}
        statusLabel={statusLabel}
        statusTone={statusTone}
      />

      <View style={styles.statsRow}>
        <View style={styles.statItem}>
          <StatCard title="Einträge" value={String(stats.entryCount)} icon="✏️" />
        </View>
        <View style={styles.statItem}>
          <StatCard title="Fotos" value={String(stats.photoCount)} icon="📷" />
        </View>
        <View style={styles.statItem}>
          <StatCard title="Offen" value={String(stats.openCount)} icon="🟠" tone="warning" />
        </View>
      </View>

      <ActionChipRow
        actions={[
          {
            key: 'export',
            label: exporting ? 'Export…' : 'Export',
            icon: '📤',
            disabled: exporting,
            onPress: () => setExportSheetVisible(true)
          },
          {
            key: 'edit',
            label: 'Bearbeiten',
            icon: '✎',
            onPress: () => router.push(`/sitereport/protocol/${protocol.id}/edit`)
          },
          {
            key: 'close',
            label: 'Abschließen',
            icon: '✓',
            disabled: exporting,
            onPress: () => setCloseVisible(true)
          }
        ]}
      />

      <View style={styles.entriesSection}>
        <SectionHeader title={`Einträge (${protocol.entries.length})`} />
        {protocol.entries.length === 0 ? (
          <EmptyState
            icon="📷"
            title="Noch keine Einträge"
            description='Tippe unten auf "+ Neuer Eintrag", um den geführten Wizard zu starten.'
          />
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
      </View>

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
  statsRow: {
    flexDirection: 'row',
    gap: spacing.sm,
    marginBottom: spacing.md
  },
  statItem: {
    flex: 1
  },
  entriesSection: {
    marginTop: spacing.lg
  }
});
