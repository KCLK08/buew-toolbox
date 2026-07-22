import { useCallback, useEffect, useMemo, useState } from 'react';
import { ActivityIndicator, Alert, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { PrimaryButton, Screen } from '../../src/components/mobile';
import { colors, spacing, typography } from '../../src/constants/theme';
import { SetupEditor } from '../../src/native/bautagebuch/components/SetupEditor';
import { getDetectedFields, saveSetupModel } from '../../src/native/bautagebuch/db/database';
import { useSetupAutosave } from '../../src/native/bautagebuch/hooks/useSetupAutosave';
import {
  ensureBuiltinTemplate,
  getTemplateBundle,
  importTemplateFromDocument,
  listManagedTemplates,
  resolveActiveTemplateId,
  setActiveTemplateId
} from '../../src/native/bautagebuch/services/templateService';
import { validateSetupModel } from '../../src/native/bautagebuch/lib/setup-model.js';
import type { BautagebuchTemplate, DetectedField } from '../../src/native/bautagebuch/types';

export default function BautagebuchSetupScreen() {
  const router = useRouter();
  const [loading, setLoading] = useState(true);
  const [templates, setTemplates] = useState<BautagebuchTemplate[]>([]);
  const [activeTemplateId, setActiveTemplateIdState] = useState('');
  const [templateId, setTemplateId] = useState('');
  const [templateName, setTemplateName] = useState('');
  const [templatePdfPath, setTemplatePdfPath] = useState<string | null>(null);
  const [detectedFields, setDetectedFields] = useState<DetectedField[]>([]);
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [importing, setImporting] = useState(false);
  const { schedule, flush } = useSetupAutosave(templateId);

  const loadEditingTemplate = useCallback(async (nextTemplateId: string) => {
    const bundle = await getTemplateBundle(nextTemplateId);
    setTemplateId(bundle.template.templateId);
    setTemplateName(bundle.template.templateName);
    setTemplatePdfPath(bundle.template.pdfPath);
    setSetupModel(bundle.setupModel);
    setDetectedFields(await getDetectedFields(bundle.template.templateId));
  }, []);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      await ensureBuiltinTemplate();
      const [templateList, activeId] = await Promise.all([
        listManagedTemplates(),
        resolveActiveTemplateId()
      ]);
      setTemplates(templateList);
      setActiveTemplateIdState(activeId);
      await loadEditingTemplate(activeId);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Setup konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, [loadEditingTemplate]);

  useEffect(() => {
    void load();
  }, [load]);

  const handleChange = (next: Record<string, unknown>) => {
    setSetupModel(next);
    setInfo('Änderungen werden gespeichert…');
    schedule(next);
  };

  const handleSelectEdit = async (nextTemplateId: string) => {
    if (nextTemplateId === templateId) return;
    setError(null);
    try {
      if (templateId) {
        await flush();
      }
      await loadEditingTemplate(nextTemplateId);
      setInfo(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Vorlage konnte nicht geladen werden.');
    }
  };

  const handleSetActive = async (nextTemplateId: string) => {
    setError(null);
    try {
      await setActiveTemplateId(nextTemplateId);
      setActiveTemplateIdState(nextTemplateId);
      setInfo('Aktive Vorlage gespeichert. Neue BTBs nutzen diese Vorlage.');
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Aktive Vorlage konnte nicht gesetzt werden.');
    }
  };

  const handleImport = async () => {
    setImporting(true);
    setError(null);
    try {
      const result = await importTemplateFromDocument();
      if (!result) return;
      const templateList = await listManagedTemplates();
      setTemplates(templateList);
      await handleSelectEdit(result.templateId);
      setInfo('PDF-Vorlage importiert. Bitte Setup prüfen und abschließen.');
    } catch (err) {
      Alert.alert('Import', err instanceof Error ? err.message : 'Vorlage konnte nicht importiert werden.');
    } finally {
      setImporting(false);
    }
  };

  const handleFinish = async () => {
    if (!templateId || !setupModel) return;
    setSaving(true);
    setError(null);
    try {
      await flush();
      const issues = validateSetupModel(setupModel);
      if (issues.length > 0) {
        setError(issues[0]);
        return;
      }
      const readyModel = {
        ...setupModel,
        status: 'ready',
        updatedAt: new Date().toISOString()
      };
      await saveSetupModel(templateId, readyModel, 'ready');
      setTemplates(await listManagedTemplates());
      setInfo('Setup abgeschlossen. Vorlage ist startbereit.');
      router.back();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Setup konnte nicht abgeschlossen werden.');
    } finally {
      setSaving(false);
    }
  };

  const validationIssues = useMemo(
    () => (setupModel ? validateSetupModel(setupModel) : []),
    [setupModel]
  );

  return (
    <Screen
      title="Setup-Editor"
      subtitle={templateName || 'Vorlagen verwalten'}
      showBack
      scroll={false}
      contentStyle={styles.screenContent}
      refreshing={loading}
      onRefresh={load}
      footer={
        setupModel ? (
          <PrimaryButton
            label={saving ? 'Wird abgeschlossen…' : 'Setup abschließen'}
            disabled={saving || validationIssues.length > 0}
            onPress={() => void handleFinish()}
          />
        ) : null
      }
    >
      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Setup wird geladen…</Text>
        </View>
      ) : null}

      {!loading && setupModel ? (
        <SetupEditor
          templates={templates}
          activeTemplateId={activeTemplateId}
          editingTemplateId={templateId}
          importing={importing}
          onSelectEdit={(id) => void handleSelectEdit(id)}
          onSetActive={(id) => void handleSetActive(id)}
          onImport={() => void handleImport()}
          templateName={templateName}
          templatePdfPath={templatePdfPath}
          detectedFields={detectedFields}
          setupModel={setupModel}
          onChange={handleChange}
          info={info}
          error={error}
        />
      ) : null}
    </Screen>
  );
}

const styles = StyleSheet.create({
  screenContent: {
    paddingHorizontal: 0,
    paddingTop: 0,
    flex: 1
  },
  center: {
    flex: 1,
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm,
    padding: spacing.xl
  },
  muted: {
    ...typography.caption,
    color: colors.muted
  }
});
