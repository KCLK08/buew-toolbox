import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ActivityIndicator,
  Alert,
  Modal,
  Pressable,
  ScrollView,
  StyleSheet,
  Text,
  View,
} from 'react-native';
import { useLocalSearchParams, useRouter } from 'expo-router';
import * as Location from 'expo-location';

import { SectionForm } from '@/components/SectionForm';
import { StatusBadge } from '@/components/StatusBadge';
import { colors } from '@/theme/colors';
import type { ExportMode, PhotoDoc, Run, SetupModel, Template } from '@/types';
import { addExportRecord, getRun, getSetupModel, getTemplate, getTemplateBytes, updateRun } from '@/lib/db';
import { sharePdfBytes } from '@/lib/export-service';
import { buildFinalPdfBytes } from '@/lib/pdf-export';
import { mergeBtbWithPhotoDoc } from '@/lib/photo-doc';
import { inputKeyForField, requiredMissingCount, sectionProgressState } from '@/lib/setup-model';
import {
  buildRunSectionsWithPhotoDoc,
  isPhotoDocEnabled,
  normalizeRunPhotoDoc,
  PHOTO_DOC_ENABLED_RUN_KEY,
  sanitizeFileName,
} from '@/lib/run-utils';
import { fetchCurrentWeatherForCoordinates, formatTemperatureValue, pickWeatherDropdownOption } from '@/lib/weather';

const WEATHER_FIELD_NAMES = {
  dropdown: 'Dropdown6',
  tempMin: 'Text11',
  tempMax: 'Text12',
};

export default function RunScreen() {
  const { id } = useLocalSearchParams<{ id: string }>();
  const router = useRouter();
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [run, setRun] = useState<Run | null>(null);
  const [template, setTemplate] = useState<Template | null>(null);
  const [model, setModel] = useState<SetupModel | null>(null);
  const [values, setValues] = useState<Record<string, string | boolean>>({});
  const [photoDoc, setPhotoDoc] = useState<PhotoDoc>({ enabled: null, entries: [], updatedAt: '' });
  const [sectionIndex, setSectionIndex] = useState(0);
  const [error, setError] = useState('');
  const [autosaveLabel, setAutosaveLabel] = useState('Bereit');
  const [weatherSyncBusy, setWeatherSyncBusy] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [exportMode, setExportMode] = useState<ExportMode>('btb_with_photo_doc');
  const saveTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  const sections = useMemo(() => (model ? buildRunSectionsWithPhotoDoc(model, photoDoc.enabled) : []), [model, photoDoc.enabled]);
  const activeSection = sections[sectionIndex];

  const sectionOptions = useMemo(() => {
    return sections.map((section) => {
      if (section.kind === 'photo-doc') {
        const missing = isPhotoDocEnabled(photoDoc, values) && photoDoc.entries.length === 0 ? 1 : 0;
        const hasAny = photoDoc.entries.length > 0;
        const state = missing > 0 ? 'progress' : hasAny ? 'done' : 'todo';
        return { section, state };
      }
      const state = sectionProgressState(section as never, values, {
        visibleRowCount: section.tableId ? Number(values[`__tableRows:${section.tableId}`] || 1) : undefined,
      });
      return { section, state };
    });
  }, [sections, values, photoDoc]);

  const loadRun = useCallback(async () => {
    if (!id) return;
    setLoading(true);
    try {
      const loadedRun = await getRun(id);
      if (!loadedRun) throw new Error('BTB nicht gefunden.');
      const loadedTemplate = await getTemplate(loadedRun.templateId);
      const loadedModel = await getSetupModel(loadedRun.templateId);
      if (!loadedTemplate || !loadedModel) throw new Error('Vorlage fehlt.');

      setRun(loadedRun);
      setTemplate(loadedTemplate);
      setModel(loadedModel);
      setValues(loadedRun.values || {});
      setPhotoDoc(normalizeRunPhotoDoc(loadedRun.photoDoc));
      setSectionIndex(Math.max(0, Math.min(loadedRun.sectionIndex || 0, buildRunSectionsWithPhotoDoc(loadedModel, loadedRun.photoDoc?.enabled).length - 1)));
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLoading(false);
    }
  }, [id]);

  useEffect(() => {
    loadRun();
  }, [loadRun]);

  function scheduleSave(nextValues = values, nextPhotoDoc = photoDoc, nextSectionIndex = sectionIndex) {
    if (!run) return;
    setAutosaveLabel('Speichert…');
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      setSaving(true);
      try {
        await updateRun(run.runId, {
          values: nextValues,
          photoDoc: nextPhotoDoc,
          sectionIndex: nextSectionIndex,
        });
        setAutosaveLabel('Gespeichert');
      } catch {
        setAutosaveLabel('Fehler');
      } finally {
        setSaving(false);
      }
    }, 500);
  }

  function handleValueChange(key: string, value: string | boolean) {
    const next = { ...values, [key]: value };
    setValues(next);
    scheduleSave(next, photoDoc, sectionIndex);
  }

  function handlePhotoDocChange(next: PhotoDoc) {
    const nextValues = { ...values, [PHOTO_DOC_ENABLED_RUN_KEY]: next.enabled === true };
    setPhotoDoc(next);
    setValues(nextValues);
    scheduleSave(nextValues, next, sectionIndex);
  }

  function goToSection(index: number) {
    setSectionIndex(index);
    scheduleSave(values, photoDoc, index);
  }

  async function syncWeather() {
    if (!model || weatherSyncBusy) return;
    setWeatherSyncBusy(true);
    try {
      const permission = await Location.requestForegroundPermissionsAsync();
      if (!permission.granted) throw new Error('Standortzugriff wird für Wetter benötigt.');
      const position = await Location.getCurrentPositionAsync({});
      const weather = await fetchCurrentWeatherForCoordinates(position.coords.latitude, position.coords.longitude);

      const weatherSection = model.single_sections.find((s) => s.sectionId === 'weather');
      const dropdownField = weatherSection?.fields.find((f) => f.fieldName === WEATHER_FIELD_NAMES.dropdown);
      const tempMinField = weatherSection?.fields.find((f) => f.fieldName === WEATHER_FIELD_NAMES.tempMin);
      const tempMaxField = weatherSection?.fields.find((f) => f.fieldName === WEATHER_FIELD_NAMES.tempMax);

      const patch: Record<string, string | boolean> = { ...values };
      if (tempMinField) patch[inputKeyForField(tempMinField)] = formatTemperatureValue(weather.tempMin);
      if (tempMaxField) patch[inputKeyForField(tempMaxField)] = formatTemperatureValue(weather.tempMax);
      if (dropdownField) {
        patch[inputKeyForField(dropdownField)] = pickWeatherDropdownOption(dropdownField.options, weather.weatherCode);
      }
      setValues(patch);
      scheduleSave(patch, photoDoc, sectionIndex);
    } catch (e) {
      Alert.alert('Wetter', (e as Error).message);
    } finally {
      setWeatherSyncBusy(false);
    }
  }

  async function exportPdf(mode: ExportMode) {
    if (!run || !template || !model) return;
    setExportOpen(false);
    try {
      const templateBytes = await getTemplateBytes(template.templateId);
      if (!templateBytes) throw new Error('PDF-Vorlage fehlt.');

      const baseBytes = await buildFinalPdfBytes({ templateBytes, setupModel: model, runValues: values });
      const baseFileName = sanitizeFileName(run.title);
      let bytes = baseBytes;
      let fileName = `${baseFileName}.pdf`;

      if (mode === 'photo_doc_only') {
        const merged = await mergeBtbWithPhotoDoc({
          btbPdfBytes: baseBytes,
          photoDocEnabled: true,
          photoEntries: photoDoc.entries,
          photoDocTitle: `Fotodokumentation - ${run.title}`,
        });
        bytes = merged.bytes;
        fileName = `${baseFileName}_Fotodoku.pdf`;
      } else if (mode === 'btb_with_photo_doc') {
        const merged = await mergeBtbWithPhotoDoc({
          btbPdfBytes: baseBytes,
          photoDocEnabled: isPhotoDocEnabled(photoDoc, values),
          photoEntries: photoDoc.entries,
          photoDocTitle: `Fotodokumentation - ${run.title}`,
        });
        bytes = merged.bytes;
      }

      await sharePdfBytes(bytes, fileName);
      await addExportRecord({ runId: run.runId, fileName });
      await updateRun(run.runId, { status: 'completed', completedAt: new Date().toISOString() });
      Alert.alert('Export', 'PDF wurde erstellt und kann geteilt werden.');
    } catch (e) {
      Alert.alert('Export', (e as Error).message);
    }
  }

  if (loading) {
    return (
      <View style={styles.centered}>
        <ActivityIndicator size="large" color={colors.primary} />
      </View>
    );
  }

  if (!run || !model || !activeSection) {
    return (
      <View style={styles.centered}>
        <Text style={styles.error}>{error || 'BTB konnte nicht geladen werden.'}</Text>
        <Pressable onPress={() => router.back()}>
          <Text style={styles.link}>Zurück</Text>
        </Pressable>
      </View>
    );
  }

  return (
    <View style={styles.container}>
      <View style={styles.header}>
        <Text style={styles.title}>{run.title}</Text>
        <Text style={styles.autosave}>{saving ? 'Speichert…' : autosaveLabel}</Text>
      </View>

      <ScrollView horizontal showsHorizontalScrollIndicator={false} style={styles.tabs} contentContainerStyle={styles.tabsContent}>
        {sectionOptions.map(({ section, state }, index) => (
          <Pressable
            key={section.sectionId}
            style={[styles.tab, index === sectionIndex && styles.tabActive]}
            onPress={() => goToSection(index)}
          >
            <Text style={[styles.tabText, index === sectionIndex && styles.tabTextActive]}>{section.label}</Text>
            <StatusBadge state={state as 'done' | 'progress' | 'todo'} />
          </Pressable>
        ))}
      </ScrollView>

      <ScrollView style={styles.form} contentContainerStyle={styles.formContent}>
        <Text style={styles.sectionTitle}>{activeSection.label}</Text>
        <SectionForm
          section={activeSection}
          values={values}
          photoDoc={photoDoc}
          onValueChange={handleValueChange}
          onPhotoDocChange={handlePhotoDocChange}
          onWeatherSync={activeSection.sectionId === 'single:weather' ? syncWeather : undefined}
          weatherSyncBusy={weatherSyncBusy}
        />
      </ScrollView>

      <View style={styles.footer}>
        <Pressable
          style={[styles.navButton, sectionIndex === 0 && styles.disabled]}
          disabled={sectionIndex === 0}
          onPress={() => goToSection(sectionIndex - 1)}
        >
          <Text style={styles.navButtonText}>Zurück</Text>
        </Pressable>
        {sectionIndex < sections.length - 1 ? (
          <Pressable style={styles.navButtonPrimary} onPress={() => goToSection(sectionIndex + 1)}>
            <Text style={styles.navButtonPrimaryText}>Weiter</Text>
          </Pressable>
        ) : (
          <Pressable style={styles.navButtonPrimary} onPress={() => setExportOpen(true)}>
            <Text style={styles.navButtonPrimaryText}>Exportieren</Text>
          </Pressable>
        )}
      </View>

      <Modal visible={exportOpen} transparent animationType="slide">
        <View style={styles.modalBackdrop}>
          <View style={styles.modalCard}>
            <Text style={styles.modalTitle}>PDF exportieren</Text>
            {([
              ['btb_only', 'Nur BTB'],
              ['photo_doc_only', 'Nur Fotodoku'],
              ['btb_with_photo_doc', 'BTB mit Fotodoku'],
            ] as const).map(([mode, label]) => (
              <Pressable
                key={mode}
                style={[styles.exportOption, exportMode === mode && styles.exportOptionActive]}
                onPress={() => setExportMode(mode)}
              >
                <Text style={styles.exportOptionText}>{label}</Text>
              </Pressable>
            ))}
            <View style={styles.modalActions}>
              <Pressable onPress={() => setExportOpen(false)}>
                <Text style={styles.cancel}>Abbrechen</Text>
              </Pressable>
              <Pressable onPress={() => exportPdf(exportMode)}>
                <Text style={styles.confirm}>Export starten</Text>
              </Pressable>
            </View>
          </View>
        </View>
      </Modal>
    </View>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: colors.background },
  centered: { flex: 1, alignItems: 'center', justifyContent: 'center', backgroundColor: colors.background, padding: 20 },
  header: { backgroundColor: colors.surface, borderBottomColor: colors.border, borderBottomWidth: 1, padding: 16 },
  title: { color: colors.text, fontSize: 18, fontWeight: '700' },
  autosave: { color: colors.textMuted, fontSize: 12, marginTop: 4 },
  tabs: { maxHeight: 88, backgroundColor: colors.surface, borderBottomColor: colors.border, borderBottomWidth: 1 },
  tabsContent: { gap: 8, padding: 12 },
  tab: { alignItems: 'center', backgroundColor: colors.background, borderRadius: 10, gap: 6, minWidth: 120, padding: 10 },
  tabActive: { backgroundColor: '#e8f4f2', borderColor: colors.primary, borderWidth: 1 },
  tabText: { color: colors.textMuted, fontSize: 12, fontWeight: '600', textAlign: 'center' },
  tabTextActive: { color: colors.primary },
  form: { flex: 1 },
  formContent: { padding: 16, paddingBottom: 32 },
  sectionTitle: { color: colors.accent, fontSize: 20, fontWeight: '800', marginBottom: 16 },
  footer: { backgroundColor: colors.surface, borderTopColor: colors.border, borderTopWidth: 1, flexDirection: 'row', gap: 10, padding: 12 },
  navButton: { borderColor: colors.border, borderRadius: 10, borderWidth: 1, flex: 1, paddingVertical: 12 },
  navButtonText: { color: colors.textMuted, fontWeight: '600', textAlign: 'center' },
  navButtonPrimary: { backgroundColor: colors.primary, borderRadius: 10, flex: 1, paddingVertical: 12 },
  navButtonPrimaryText: { color: '#fff', fontWeight: '700', textAlign: 'center' },
  disabled: { opacity: 0.4 },
  error: { color: colors.danger, marginBottom: 12, textAlign: 'center' },
  link: { color: colors.primary, fontWeight: '600' },
  modalBackdrop: { backgroundColor: 'rgba(0,0,0,0.45)', flex: 1, justifyContent: 'flex-end' },
  modalCard: { backgroundColor: colors.surface, borderTopLeftRadius: 16, borderTopRightRadius: 16, padding: 20 },
  modalTitle: { color: colors.text, fontSize: 18, fontWeight: '700', marginBottom: 12 },
  exportOption: { borderColor: colors.border, borderRadius: 10, borderWidth: 1, marginBottom: 8, padding: 12 },
  exportOptionActive: { backgroundColor: '#e8f4f2', borderColor: colors.primary },
  exportOptionText: { color: colors.text, fontWeight: '600' },
  modalActions: { flexDirection: 'row', justifyContent: 'space-between', marginTop: 12 },
  cancel: { color: colors.textMuted, fontWeight: '600' },
  confirm: { color: colors.primary, fontWeight: '700' },
});
