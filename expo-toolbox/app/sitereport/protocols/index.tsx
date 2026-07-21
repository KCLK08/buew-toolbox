import { useCallback, useEffect, useState } from 'react';
import { Alert, Pressable, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { BottomSheet, ProtocolCard } from '../../../src/components/sitereport';
import { EmptyState, PrimaryButton, Screen } from '../../../src/components/mobile';
import { useToast } from '../../../src/contexts/ToastContext';
import { colors, spacing, typography } from '../../../src/constants/theme';
import {
  initSiteReportDatabase,
  listProtocols,
  type SiteReportProtocol
} from '../../../src/native/sitereport/db/database';
import {
  bulkDeleteProtocols,
  deleteProtocolWithCleanup,
  protocolStats
} from '../../../src/native/sitereport/services/protocolService';
import {
  closeProtocolWithExport,
  exportProtocolPdf,
  exportProtocolXlsx,
  type CloseExportMode
} from '../../../src/native/sitereport/services/exportService';

export default function ProtocolsListScreen() {
  const router = useRouter();
  const { showToast } = useToast();
  const [loading, setLoading] = useState(true);
  const [protocols, setProtocols] = useState<SiteReportProtocol[]>([]);
  const [selectionMode, setSelectionMode] = useState(false);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [exportSheetVisible, setExportSheetVisible] = useState(false);
  const [busy, setBusy] = useState(false);

  const load = useCallback(async () => {
    setLoading(true);
    await initSiteReportDatabase();
    setProtocols(await listProtocols());
    setLoading(false);
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const toggleSelect = (id: string) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const selectAll = () => setSelected(new Set(protocols.map((p) => p.id)));

  const deleteSelected = () => {
    if (selected.size === 0) return;
    Alert.alert('Löschen', `${selected.size} Protokoll(e) wirklich löschen?`, [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Löschen',
        style: 'destructive',
        onPress: () => {
          void bulkDeleteProtocols([...selected]).then(() => {
            showToast('Protokolle gelöscht');
            setSelected(new Set());
            setSelectionMode(false);
            void load();
          });
        }
      }
    ]);
  };

  const exportSelected = async (mode: CloseExportMode) => {
    setExportSheetVisible(false);
    setBusy(true);
    try {
      const targets = protocols.filter((p) => selected.has(p.id));
      for (const protocol of targets) {
        if (mode === 'pdf' || mode === 'both') {
          await exportProtocolPdf(protocol, { share: false });
        }
        if (mode === 'xlsx' || mode === 'both') {
          await exportProtocolXlsx(protocol, { share: false });
        }
        if (mode === 'save') {
          await closeProtocolWithExport(protocol, 'save');
        }
      }
      showToast('Export abgeschlossen');
    } catch (err) {
      Alert.alert('Export', err instanceof Error ? err.message : 'Export fehlgeschlagen.');
    } finally {
      setBusy(false);
    }
  };

  return (
    <Screen title="Protokolle" subtitle={`${protocols.length} gespeichert`} showBack scroll refreshing={loading} onRefresh={load}>
      {protocols.length === 0 ? (
        <EmptyState title="Keine Protokolle" description="Starte ein neues Protokoll vom Dashboard." />
      ) : (
        protocols.map((protocol) => {
          const stats = protocolStats(protocol);
          return (
            <ProtocolCard
              key={protocol.id}
              title={protocol.protocolTitle}
              subtitle={protocol.projectName}
              meta={`${protocol.protocolDate} · ${stats.entryCount} Einträge · ${stats.photoCount} Fotos`}
              selected={selected.has(protocol.id)}
              onSelectToggle={selectionMode ? () => toggleSelect(protocol.id) : undefined}
              onPress={
                selectionMode
                  ? () => toggleSelect(protocol.id)
                  : () => router.push(`/sitereport/protocol/${protocol.id}`)
              }
              trailing={
                !selectionMode ? (
                  <Pressable
                    hitSlop={8}
                    onPress={() => {
                      Alert.alert('Löschen', 'Protokoll wirklich löschen?', [
                        { text: 'Abbrechen', style: 'cancel' },
                        {
                          text: 'Löschen',
                          style: 'destructive',
                          onPress: () => {
                            void deleteProtocolWithCleanup(protocol.id).then(() => {
                              showToast('Protokoll gelöscht');
                              void load();
                            });
                          }
                        }
                      ]);
                    }}
                  >
                    <Text style={styles.delete}>✕</Text>
                  </Pressable>
                ) : undefined
              }
            />
          );
        })
      )}

      {protocols.length > 0 ? (
        <View style={styles.toolbar}>
          <PrimaryButton
            label={selectionMode ? 'Fertig' : 'Auswählen'}
            variant="secondary"
            onPress={() => {
              setSelectionMode((prev) => !prev);
              setSelected(new Set());
            }}
          />
          {selectionMode ? (
            <>
              <PrimaryButton label="Alle" variant="ghost" onPress={selectAll} />
              <PrimaryButton
                label="Export"
                variant="secondary"
                disabled={selected.size === 0 || busy}
                onPress={() => setExportSheetVisible(true)}
              />
              <PrimaryButton
                label="Löschen"
                variant="ghost"
                disabled={selected.size === 0}
                onPress={deleteSelected}
              />
            </>
          ) : null}
        </View>
      ) : null}

      <BottomSheet visible={exportSheetVisible} title="Auswahl exportieren" onClose={() => setExportSheetVisible(false)}>
        <PrimaryButton label="PDF" onPress={() => void exportSelected('pdf')} disabled={busy} />
        <PrimaryButton label="Excel" variant="secondary" onPress={() => void exportSelected('xlsx')} disabled={busy} />
        <PrimaryButton label="PDF + Excel" variant="secondary" onPress={() => void exportSelected('both')} disabled={busy} />
      </BottomSheet>
    </Screen>
  );
}

const styles = StyleSheet.create({
  toolbar: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs,
    marginTop: spacing.md
  },
  delete: {
    color: colors.danger,
    fontSize: 18,
    padding: 4
  }
});
