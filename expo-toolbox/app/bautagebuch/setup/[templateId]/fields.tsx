import { useCallback, useEffect, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter, type Href } from 'expo-router';
import { SafeAreaView } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../src/constants/theme';
import { SetupEditor } from '../../../../src/native/bautagebuch/components/SetupEditor';
import { SetupFieldsStep } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupFieldsStep';
import { SetupWizardStepNav } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupWizardStepNav';
import { getDetectedFields, saveSetupModel, updateTemplateField } from '../../../../src/native/bautagebuch/db/database';
import { useSetupAutosave } from '../../../../src/native/bautagebuch/hooks/useSetupAutosave';
import {
  hasTableSections,
  getWizardState,
  markSetupCompleted,
  rebuildSectionsFromWizard,
  resolveSetupEditStepPath,
  resolveSetupEntryPath,
  resolveTemplateDetailPath,
  resolveTemplateEditPath,
  shouldShowFieldsIntro,
  sortMappingFields
} from '../../../../src/native/bautagebuch/lib/setup-mapping';
import { validateSetupModel } from '../../../../src/native/bautagebuch/lib/setup-model.js';
import { getTemplateBundle } from '../../../../src/native/bautagebuch/services/templateService';

export default function SetupFieldsScreen() {
  const router = useRouter();
  const { templateId } = useLocalSearchParams<{ templateId: string }>();
  const [loading, setLoading] = useState(true);
  const [templateName, setTemplateName] = useState('');
  const [templateKind, setTemplateKind] = useState('');
  const [templateStatus, setTemplateStatus] = useState('');
  const [pdfPath, setPdfPath] = useState<string | null>(null);
  const [detectedFields, setDetectedFields] = useState<Awaited<ReturnType<typeof getDetectedFields>>>([]);
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const { schedule, flush } = useSetupAutosave(String(templateId || ''));

  const load = useCallback(async () => {
    if (!templateId) return;
    setLoading(true);
    setError(null);
    try {
      const bundle = await getTemplateBundle(templateId);
      setTemplateName(bundle.template.templateName);
      setTemplateKind(bundle.template.templateKind);
      setTemplateStatus(bundle.template.status);
      setPdfPath(bundle.template.pdfPath);
      const fields = await getDetectedFields(templateId);
      setDetectedFields(fields);
      const mappingFields = sortMappingFields(fields);
      let nextModel = bundle.setupModel;
      const wizard = getWizardState(nextModel);
      if (wizard.editMode && wizard.step === 'fields') {
        nextModel = rebuildSectionsFromWizard(nextModel, mappingFields);
      }
      setSetupModel(nextModel);
      if (nextModel !== bundle.setupModel) {
        schedule(nextModel);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Setup konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, [templateId, schedule]);

  const reloadFields = useCallback(async () => {
    if (!templateId) return;
    setDetectedFields(await getDetectedFields(templateId));
  }, [templateId]);

  const handleUpdateField = useCallback(
    async (fieldId: string, patch: { labelCandidate?: string; type?: string }) => {
      if (!templateId) return;
      await updateTemplateField(templateId, fieldId, patch);
      await reloadFields();
    },
    [templateId, reloadFields]
  );

  useEffect(() => {
    void load();
  }, [load]);

  useEffect(() => {
    if (loading || !setupModel || !templateId) return;
    const wizard = getWizardState(setupModel);
    const legacy =
      templateKind === 'builtin-etb' || (hasTableSections(setupModel) && wizard.step !== 'fields');
    if (legacy) return;
    if (wizard.step !== 'fields') {
      router.replace(resolveSetupEntryPath(String(templateId), setupModel, templateKind));
      return;
    }
    if (shouldShowFieldsIntro(setupModel, templateKind)) {
      router.replace(`/bautagebuch/setup/${templateId}/fields-intro` as Href);
    }
  }, [loading, setupModel, templateKind, templateId, router]);

  const wizard = setupModel ? getWizardState(setupModel) : null;
  const useLegacyEditor = Boolean(
    setupModel &&
      (templateKind === 'builtin-etb' ||
        (hasTableSections(setupModel) && wizard?.step !== 'fields'))
  );
  const readOnly = templateStatus === 'archived';
  const editMode = wizard?.editMode === true;

  const handleChange = (next: Record<string, unknown>) => {
    setSetupModel(next);
    schedule(next);
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
      const readyModel = markSetupCompleted({
        ...setupModel,
        status: 'ready',
        updatedAt: new Date().toISOString()
      });
      await saveSetupModel(templateId, readyModel, 'ready');
      setSetupModel(readyModel);
      router.replace(resolveTemplateDetailPath(String(templateId)));
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Setup konnte nicht abgeschlossen werden.');
    } finally {
      setSaving(false);
    }
  };

  const handleBack = async () => {
    await flush();
    if (editMode) {
      router.replace(resolveTemplateEditPath(String(templateId)));
      return;
    }
    router.replace(`/bautagebuch/setup/${templateId}/assign` as Href);
  };

  const handleStepNav = async (step: 'structure' | 'assign' | 'fields') => {
    await flush();
    router.replace(resolveSetupEditStepPath(String(templateId), step));
  };

  return (
    <SafeAreaView style={styles.root} edges={['top', 'left', 'right']}>
      {editMode ? (
        <SetupWizardStepNav activeStep="fields" onSelectStep={(step) => void handleStepNav(step)} />
      ) : null}

      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Feldeinstellungen werden geladen…</Text>
        </View>
      ) : null}

      {error ? <Text style={styles.error}>{error}</Text> : null}

      {!loading && setupModel && useLegacyEditor ? (
        <SetupEditor
          templates={[]}
          activeTemplateId={String(templateId)}
          editingTemplateId={String(templateId)}
          templateName={templateName}
          templatePdfPath={pdfPath}
          detectedFields={detectedFields}
          setupModel={setupModel}
          onChange={handleChange}
          error={error}
          embedded
          onTemplateRenamed={setTemplateName}
          onSelectEdit={() => undefined}
          onSetActive={() => undefined}
          onImport={() => undefined}
        />
      ) : null}

      {!loading && setupModel && !useLegacyEditor ? (
        <SetupFieldsStep
          pdfPath={pdfPath}
          detectedFields={detectedFields}
          setupModel={setupModel}
          readOnly={readOnly}
          onChange={handleChange}
          onFinish={() => void handleFinish()}
          onBack={() => void handleBack()}
          onUpdateField={handleUpdateField}
        />
      ) : null}

      {saving ? (
        <View style={styles.savingOverlay}>
          <ActivityIndicator color={colors.accent} />
        </View>
      ) : null}
    </SafeAreaView>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: colors.bg
  },
  center: {
    flex: 1,
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm
  },
  muted: {
    ...typography.caption,
    color: colors.muted
  },
  error: {
    ...typography.caption,
    color: colors.danger,
    padding: spacing.pageX
  },
  savingOverlay: {
    ...StyleSheet.absoluteFillObject,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: 'rgba(242, 240, 235, 0.6)'
  }
});
