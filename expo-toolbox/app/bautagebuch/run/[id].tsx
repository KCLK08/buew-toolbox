import { useCallback, useEffect, useState } from 'react';
import { Alert, StyleSheet, Text } from 'react-native';
import { useLocalSearchParams } from 'expo-router';

import { PrimaryButton, Screen } from '../../../src/components/mobile';
import { colors, typography } from '../../../src/constants/theme';
import { RunWizard } from '../../../src/native/bautagebuch/components/RunWizard';
import { getRun, updateRun } from '../../../src/native/bautagebuch/db/database';
import {
  exportRunPdf,
  type BautagebuchExportMode
} from '../../../src/native/bautagebuch/services/exportService';
import { capturePhotoDocEntry } from '../../../src/native/bautagebuch/services/photoDocService';
import { getActiveTemplateBundle } from '../../../src/native/bautagebuch/services/templateService';
import { syncWeatherValues } from '../../../src/native/bautagebuch/services/weatherService';
import type { BautagebuchRun } from '../../../src/native/bautagebuch/types';
import { nowIso } from '../../../src/lib/ids';

export default function BautagebuchRunScreen() {
  const { id } = useLocalSearchParams<{ id: string }>();
  const [run, setRun] = useState<BautagebuchRun | null>(null);
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [exporting, setExporting] = useState(false);
  const [weatherBusy, setWeatherBusy] = useState(false);

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

  const persist = async (patch: Partial<BautagebuchRun>) => {
    if (!run) return;
    const next = await updateRun(run.runId, patch);
    if (next) setRun(next);
  };

  const handleWeatherSync = async () => {
    if (!run || !setupModel) return;
    setWeatherBusy(true);
    try {
      const weather = await syncWeatherValues();
      const sections =
        (setupModel.single_sections as Array<{
          sectionId: string;
          fields: Array<{ fieldName: string; fieldId: string }>;
        }>) || [];
      const weatherSection = sections.find((section) => section.sectionId === 'weather');
      const nextValues = { ...run.values };
      const setByName = (name: string, value: string) => {
        const field = weatherSection?.fields.find((entry) => entry.fieldName === name);
        if (field) nextValues[`field:${field.fieldId}`] = value;
      };
      setByName('Dropdown6', weather.weather);
      setByName('Text11', weather.tempMin);
      setByName('Text12', weather.tempMax);
      await persist({ values: nextValues });
    } catch (err) {
      Alert.alert('Wetter', err instanceof Error ? err.message : 'Wetter konnte nicht geladen werden.');
    } finally {
      setWeatherBusy(false);
    }
  };

  const runExport = async (mode: BautagebuchExportMode) => {
    if (!run) return;
    setExporting(true);
    try {
      await exportRunPdf(run.runId, mode);
      await persist({ status: 'completed', completedAt: new Date().toISOString() });
      Alert.alert('Export', 'PDF wurde erstellt und kann geteilt werden.');
    } catch (err) {
      Alert.alert('Export', err instanceof Error ? err.message : 'PDF-Export fehlgeschlagen.');
    } finally {
      setExporting(false);
    }
  };

  const handleExport = () => {
    Alert.alert('PDF exportieren', 'Welche Version soll erstellt werden?', [
      { text: 'Abbrechen', style: 'cancel' },
      { text: 'Nur BTB', onPress: () => void runExport('btb') },
      { text: 'Nur Fotodoku', onPress: () => void runExport('photo') },
      { text: 'Zusammengeführt', onPress: () => void runExport('merged') }
    ]);
  };

  const handlePhotoDocChange = (enabled: boolean) => {
    if (!run) return;
    void persist({
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
      await persist({
        photoDoc: {
          enabled: true,
          entries: [entry, ...(run.photoDoc?.entries || [])],
          updatedAt: nowIso()
        }
      });
    } catch (err) {
      Alert.alert('Foto', err instanceof Error ? err.message : 'Foto konnte nicht gespeichert werden.');
    }
  };

  const handleRemovePhoto = (entryId: string) => {
    if (!run) return;
    void persist({
      photoDoc: {
        ...run.photoDoc,
        entries: (run.photoDoc?.entries || []).filter((entry) => entry.id !== entryId),
        updatedAt: nowIso()
      }
    });
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
        <PrimaryButton
          label={exporting ? 'PDF wird erstellt…' : 'PDF exportieren'}
          disabled={exporting}
          onPress={handleExport}
        />
      }
    >
      <RunWizard
        setupModel={setupModel}
        values={run.values}
        sectionIndex={run.sectionIndex}
        weatherBusy={weatherBusy}
        photoDoc={run.photoDoc}
        onPhotoDocChange={handlePhotoDocChange}
        onAddPhoto={() => void handleAddPhoto()}
        onRemovePhoto={handleRemovePhoto}
        onWeatherSync={handleWeatherSync}
        onSectionChange={(sectionIndex) => void persist({ sectionIndex })}
        onChange={(values) => void persist({ values })}
      />
    </Screen>
  );
}

const styles = StyleSheet.create({
  error: { ...typography.body, color: colors.danger }
});
