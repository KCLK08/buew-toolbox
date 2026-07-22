import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Alert, Pressable, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams } from 'expo-router';

import { PrimaryButton, Screen, StatusBadge } from '../../../src/components/mobile';
import { colors, typography } from '../../../src/constants/theme';
import { useToast } from '../../../src/contexts/ToastContext';
import { PdfPreviewPanel } from '../../../src/native/bautagebuch/components/PdfPreviewPanel';
import { RunWizard } from '../../../src/native/bautagebuch/components/RunWizard';
import { getRun, updateRun } from '../../../src/native/bautagebuch/db/database';
import { useRunAutosave } from '../../../src/native/bautagebuch/hooks/useRunAutosave';
import {
  computeTotalMissingRequired,
  exportBlockedMessage
} from '../../../src/native/bautagebuch/lib/run-validation';
import {
  exportRunPdf,
  generateRunPreviewPdfPath,
  type BautagebuchExportMode
} from '../../../src/native/bautagebuch/services/exportService';
import {
  capturePhotoDocEntry,
  pickPhotoDocEntry,
  removePhotoDocEntry
} from '../../../src/native/bautagebuch/services/photoDocService';
import { getActiveTemplateBundle } from '../../../src/native/bautagebuch/services/templateService';
import { syncWeatherValues } from '../../../src/native/bautagebuch/services/weatherService';
import type { BautagebuchRun } from '../../../src/native/bautagebuch/types';
import { nowIso } from '../../../src/lib/ids';

const PREVIEW_DEBOUNCE_MS = 900;

export default function BautagebuchRunScreen() {
  const { id } = useLocalSearchParams<{ id: string }>();
  const { showToast } = useToast();
  const [run, setRun] = useState<BautagebuchRun | null>(null);
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [exporting, setExporting] = useState(false);
  const [weatherBusy, setWeatherBusy] = useState(false);
  const [showPreview, setShowPreview] = useState(true);
  const [previewPath, setPreviewPath] = useState<string | null>(null);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewError, setPreviewError] = useState<string | null>(null);
  const previewTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const { schedule, flush } = useRunAutosave(id);

  const load = useCallback(async () => {
    if (!id) return;
    try {
      const [loadedRun, bundle] = await Promise.all([getRun(id), getActiveTemplateBundle()]);
      if (!loadedRun) {
        setError('BTB-Lauf nicht gefunden.');
        return;
      }
      setRun(loadedRun);
      setSetupModel(bundle.setupModel);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Laden fehlgeschlagen.');
    }
  }, [id]);

  useEffect(() => {
    void load();
  }, [load]);

  const totalMissingRequired = useMemo(() => {
    if (!setupModel || !run) return 0;
    return computeTotalMissingRequired(setupModel, run.values, run.photoDoc?.enabled ?? null);
  }, [run, setupModel]);

  const refreshPreview = useCallback(async () => {
    if (!run || !showPreview) return;
    setPreviewLoading(true);
    setPreviewError(null);
    try {
      await flush();
      const path = await generateRunPreviewPdfPath(run.runId);
      setPreviewPath(path);
    } catch (err) {
      setPreviewError(err instanceof Error ? err.message : 'Vorschau fehlgeschlagen.');
    } finally {
      setPreviewLoading(false);
    }
  }, [flush, run, showPreview]);

  useEffect(() => {
    if (!run || !showPreview) return undefined;
    if (previewTimerRef.current) clearTimeout(previewTimerRef.current);
    previewTimerRef.current = setTimeout(() => {
      void refreshPreview();
    }, PREVIEW_DEBOUNCE_MS);
    return () => {
      if (previewTimerRef.current) clearTimeout(previewTimerRef.current);
    };
  }, [run?.values, run?.photoDoc, showPreview, refreshPreview]);

  const persist = (patch: Partial<BautagebuchRun>) => {
    if (!run) return;
    const next = { ...run, ...patch };
    setRun(next);
    schedule(patch);
  };

  const handleWeatherSync = async () => {
    if (!run || !setupModel) return;
    setWeatherBusy(true);
    try {
      const sections =
        (setupModel.single_sections as Array<{
          sectionId: string;
          fields: Array<{ fieldName: string; fieldId: string; options?: string[] }>;
        }>) || [];
      const weatherSection = sections.find((section) => section.sectionId === 'weather');
      const dropdownField = weatherSection?.fields.find((field) => field.fieldName === 'Dropdown6');
      const weather = await syncWeatherValues(dropdownField?.options || []);
      const nextValues = { ...run.values };
      const setByName = (name: string, value: string) => {
        const field = weatherSection?.fields.find((entry) => entry.fieldName === name);
        if (field) nextValues[`field:${field.fieldId}`] = value;
      };
      setByName('Dropdown6', weather.weather);
      setByName('Text11', weather.tempMin);
      setByName('Text12', weather.tempMax);
      persist({ values: nextValues });
      showToast('Wetter aktualisiert');
    } catch (err) {
      Alert.alert('Wetter', err instanceof Error ? err.message : 'Wetter konnte nicht geladen werden.');
    } finally {
      setWeatherBusy(false);
    }
  };

  const runExport = async (mode: BautagebuchExportMode) => {
    if (!run) return;
    if (totalMissingRequired > 0) {
      Alert.alert('Export blockiert', exportBlockedMessage(totalMissingRequired));
      return;
    }
    setExporting(true);
    try {
      await flush();
      await exportRunPdf(run.runId, mode);
      const completed = await updateRun(run.runId, {
        status: 'completed',
        completedAt: nowIso()
      });
      if (completed) setRun(completed);
      showToast('PDF exportiert');
    } catch (err) {
      Alert.alert('Export', err instanceof Error ? err.message : 'PDF-Export fehlgeschlagen.');
    } finally {
      setExporting(false);
    }
  };

  const handleExport = () => {
    if (totalMissingRequired > 0) {
      Alert.alert('Export blockiert', exportBlockedMessage(totalMissingRequired));
      return;
    }
    Alert.alert('PDF exportieren', 'Welche Version soll erstellt werden?', [
      { text: 'Abbrechen', style: 'cancel' },
      { text: 'Nur BTB', onPress: () => void runExport('btb') },
      { text: 'Nur Fotodoku', onPress: () => void runExport('photo') },
      { text: 'Zusammengeführt', onPress: () => void runExport('merged') }
    ]);
  };

  const handlePhotoDocChange = (enabled: boolean) => {
    if (!run) return;
    persist({
      photoDoc: {
        ...run.photoDoc,
        enabled,
        updatedAt: nowIso()
      }
    });
  };

  const handleAddPhoto = async () => {
    if (!run) return;
    try {
      const entry = await capturePhotoDocEntry(run.runId);
      if (!entry) return;
      persist({
        photoDoc: {
          enabled: true,
          entries: [entry, ...(run.photoDoc?.entries || [])],
          updatedAt: nowIso()
        }
      });
      showToast('Foto gespeichert');
    } catch (err) {
      Alert.alert('Foto', err instanceof Error ? err.message : 'Foto konnte nicht gespeichert werden.');
    }
  };

  const handlePickPhoto = async () => {
    if (!run) return;
    try {
      const entry = await pickPhotoDocEntry(run.runId);
      if (!entry) return;
      persist({
        photoDoc: {
          enabled: true,
          entries: [entry, ...(run.photoDoc?.entries || [])],
          updatedAt: nowIso()
        }
      });
      showToast('Foto hinzugefügt');
    } catch (err) {
      Alert.alert('Foto', err instanceof Error ? err.message : 'Foto konnte nicht gespeichert werden.');
    }
  };

  const handleRemovePhoto = (entryId: string) => {
    if (!run) return;
    const entry = (run.photoDoc?.entries || []).find((item) => item.id === entryId);
    Alert.alert('Foto entfernen', 'Dieses Foto wirklich löschen?', [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Löschen',
        style: 'destructive',
        onPress: () => {
          void removePhotoDocEntry(run.runId, entryId, entry?.localPath).then(() => {
            persist({
              photoDoc: {
                ...run.photoDoc,
                entries: (run.photoDoc?.entries || []).filter((item) => item.id !== entryId),
                updatedAt: nowIso()
              }
            });
            showToast('Foto entfernt');
          });
        }
      }
    ]);
  };

  if (!run || !setupModel) {
    return (
      <Screen title="Bautagebuch" showBack>
        <Text style={styles.error}>{error || 'Wird geladen…'}</Text>
      </Screen>
    );
  }

  return (
    <Screen
      title={run.title}
      subtitle="Guided Flow · PDF-Export"
      showBack
      footer={
        <View style={styles.footerCol}>
          {totalMissingRequired > 0 ? (
            <StatusBadge
              label={`${totalMissingRequired} Pflichtfeld${totalMissingRequired === 1 ? '' : 'er'} offen`}
              tone="warning"
            />
          ) : (
            <StatusBadge label="Bereit zum Export" tone="success" />
          )}
          <Pressable onPress={() => setShowPreview((value) => !value)}>
            <Text style={styles.previewToggle}>
              {showPreview ? 'Vorschau ausblenden' : 'Live-PDF-Vorschau anzeigen'}
            </Text>
          </Pressable>
          <View style={styles.footerRow}>
            <PrimaryButton
              label={exporting ? 'PDF…' : 'PDF exportieren'}
              disabled={exporting || totalMissingRequired > 0}
              onPress={handleExport}
            />
          </View>
        </View>
      }
    >
      <RunWizard
        setupModel={setupModel}
        values={run.values}
        sectionIndex={run.sectionIndex}
        weatherBusy={weatherBusy}
        photoDoc={run.photoDoc}
        totalMissingRequired={totalMissingRequired}
        showPreview={showPreview}
        previewPanel={
          <PdfPreviewPanel pdfPath={previewPath} loading={previewLoading} error={previewError} />
        }
        onPhotoDocChange={handlePhotoDocChange}
        onAddPhoto={() => void handleAddPhoto()}
        onPickPhoto={() => void handlePickPhoto()}
        onRemovePhoto={handleRemovePhoto}
        onWeatherSync={handleWeatherSync}
        onSectionChange={(sectionIndex) => persist({ sectionIndex })}
        onChange={(values) => persist({ values })}
        onRequestExport={handleExport}
      />
    </Screen>
  );
}

const styles = StyleSheet.create({
  error: { ...typography.body, color: colors.danger },
  footerCol: { gap: 8 },
  footerRow: { flexDirection: 'row', gap: 8 },
  previewToggle: { ...typography.caption, color: colors.accent, textAlign: 'center' }
});
