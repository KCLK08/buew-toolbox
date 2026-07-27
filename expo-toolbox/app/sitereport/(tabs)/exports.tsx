import { useCallback, useEffect, useState } from 'react';
import { Alert, StyleSheet, View } from 'react-native';

import { ExportCard, SectionHeader } from '../../../src/components/sitereport';
import { EmptyState, PrimaryButton, Screen } from '../../../src/components/mobile';
import { useToast } from '../../../src/contexts/ToastContext';
import { spacing } from '../../../src/constants/theme';
import { hapticSuccess } from '../../../src/lib/haptics';
import { initSiteReportDatabase, type SiteReportExport } from '../../../src/native/sitereport/db/database';
import { deleteCachedExport, listExports, shareCachedExport } from '../../../src/native/sitereport/services/exportService';

export default function ExportsCenterScreen() {
  const { showToast } = useToast();
  const [loading, setLoading] = useState(true);
  const [exportsList, setExportsList] = useState<SiteReportExport[]>([]);
  const [selectionMode, setSelectionMode] = useState(false);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [sharing, setSharing] = useState<string | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    await initSiteReportDatabase();
    setExportsList(await listExports());
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

  const share = async (exportId: string, format: 'pdf' | 'xlsx') => {
    setSharing(`${exportId}:${format}`);
    try {
      await shareCachedExport(exportId, format);
      void hapticSuccess();
      showToast('Export geteilt');
    } catch (err) {
      Alert.alert('Export', err instanceof Error ? err.message : 'Teilen fehlgeschlagen.');
    } finally {
      setSharing(null);
    }
  };

  const remove = (exportId: string) => {
    Alert.alert('Export löschen', 'Gespeicherten Export wirklich entfernen?', [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Löschen',
        style: 'destructive',
        onPress: () => {
          void deleteCachedExport(exportId).then(() => {
            showToast('Export gelöscht');
            void load();
          });
        }
      }
    ]);
  };

  const deleteSelected = () => {
    if (selected.size === 0) return;
    Alert.alert('Löschen', `${selected.size} Export(e) wirklich löschen?`, [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Löschen',
        style: 'destructive',
        onPress: () => {
          void Promise.all([...selected].map((id) => deleteCachedExport(id))).then(() => {
            showToast('Exporte gelöscht');
            setSelected(new Set());
            setSelectionMode(false);
            void load();
          });
        }
      }
    ]);
  };

  return (
    <Screen
      title="Exporte"
      subtitle="Export-Center"
      toolboxBack
      reserveTabBarSpace
      scroll
      refreshing={loading}
      onRefresh={load}
    >
      {exportsList.length > 0 ? (
        <SectionHeader
          title={`${exportsList.length} Export${exportsList.length === 1 ? '' : 'e'}`}
          actionLabel={selectionMode ? 'Fertig' : 'Auswahl'}
          onAction={() => {
            setSelectionMode((prev) => !prev);
            setSelected(new Set());
          }}
        />
      ) : null}

      {exportsList.length === 0 ? (
        <EmptyState
          icon="📤"
          title="Keine Exporte"
          description="Exporte erscheinen hier nach dem Abschluss eines Protokolls."
        />
      ) : (
        exportsList.map((item) => (
          <ExportCard
            key={item.id}
            item={item}
            selected={selected.has(item.id)}
            onSelectToggle={selectionMode ? () => toggleSelect(item.id) : undefined}
            sharing={Boolean(sharing)}
            onSharePdf={
              item.pdfPath || item.pdfFilename
                ? () => void share(item.id, 'pdf')
                : undefined
            }
            onShareXlsx={
              item.xlsxPath || item.xlsxFilename
                ? () => void share(item.id, 'xlsx')
                : undefined
            }
            onDelete={selectionMode ? undefined : () => remove(item.id)}
          />
        ))
      )}

      {selectionMode && selected.size > 0 ? (
        <View style={styles.toolbar}>
          <PrimaryButton label={`${selected.size} löschen`} variant="ghost" onPress={deleteSelected} />
        </View>
      ) : null}
    </Screen>
  );
}

const styles = StyleSheet.create({
  toolbar: {
    marginTop: spacing.md
  }
});
