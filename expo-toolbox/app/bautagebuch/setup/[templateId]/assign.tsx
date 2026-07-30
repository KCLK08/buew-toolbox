import { useCallback, useEffect, useMemo, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter, type Href } from 'expo-router';
import { SafeAreaView } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../src/constants/theme';
import { SetupAssignStep } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupAssignStep';
import { SetupWizardStepNav } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupWizardStepNav';
import { addTemplateField, deleteTemplateField, getDetectedFields, updateTemplateField } from '../../../../src/native/bautagebuch/db/database';
import { useSetupAutosave } from '../../../../src/native/bautagebuch/hooks/useSetupAutosave';
import {
  ensureWizardInitialized,
  getWizardState,
  rebuildSectionsFromWizard,
  resolveFieldsSetupPath,
  resolveSetupEditStepPath,
  resolveTemplateEditPath,
  shouldShowAssignIntro,
  sortMappingFields
} from '../../../../src/native/bautagebuch/lib/setup-mapping';
import { getTemplateBundle } from '../../../../src/native/bautagebuch/services/templateService';

export default function SetupAssignScreen() {
  const router = useRouter();
  const { templateId } = useLocalSearchParams<{ templateId: string }>();
  const [loading, setLoading] = useState(true);
  const [templateStatus, setTemplateStatus] = useState('');
  const [pdfPath, setPdfPath] = useState<string | null>(null);
  const [detectedFields, setDetectedFields] = useState<Awaited<ReturnType<typeof getDetectedFields>>>([]);
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [error, setError] = useState<string | null>(null);
  const { schedule, flush } = useSetupAutosave(String(templateId || ''));

  const load = useCallback(async () => {
    if (!templateId) return;
    setLoading(true);
    setError(null);
    try {
      const bundle = await getTemplateBundle(templateId);
      const fields = await getDetectedFields(templateId);
      const sortedFields = sortMappingFields(fields);
      const wizard = getWizardState(bundle.setupModel);

      if (wizard.step === 'fields') {
        setLoading(false);
        router.replace(resolveFieldsSetupPath(String(templateId), bundle.setupModel));
        return;
      }
      if (wizard.step === 'structure') {
        setLoading(false);
        router.replace(`/bautagebuch/setup/${templateId}/mapping`);
        return;
      }

      const initialized = ensureWizardInitialized(bundle.setupModel);
      if (shouldShowAssignIntro(initialized)) {
        setLoading(false);
        router.replace(`/bautagebuch/setup/${templateId}/assign-intro` as Href);
        return;
      }

      setTemplateStatus(bundle.template.status);
      setPdfPath(bundle.template.pdfPath);
      setDetectedFields(fields);
      setSetupModel(initialized);

      if (initialized !== bundle.setupModel) {
        schedule(initialized);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Feldzuordnung konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, [templateId, schedule, flush, router]);

  useEffect(() => {
    void load();
  }, [load]);

  const mappingFields = useMemo(() => sortMappingFields(detectedFields), [detectedFields]);

  const handleChange = (next: Record<string, unknown>) => {
    setSetupModel(next);
    schedule(next);
  };

  const handleComplete = (sourceModel: Record<string, unknown>) => {
    const rebuilt = rebuildSectionsFromWizard(sourceModel, mappingFields);
    setSetupModel(rebuilt);
    schedule(rebuilt);
    void (async () => {
      await flush();
      router.replace(resolveFieldsSetupPath(String(templateId), rebuilt));
    })();
  };

  const handleBack = async () => {
    await flush();
    const wizard = setupModel ? getWizardState(setupModel) : null;
    if (wizard?.editMode) {
      router.replace(resolveTemplateEditPath(String(templateId)));
      return;
    }
    router.replace(`/bautagebuch/setup/${templateId}/mapping`);
  };

  const editMode = setupModel ? getWizardState(setupModel).editMode === true : false;

  const handleStepNav = async (step: 'structure' | 'assign' | 'fields') => {
    await flush();
    router.replace(resolveSetupEditStepPath(String(templateId), step));
  };

  const reloadFields = useCallback(async () => {
    if (!templateId) return;
    setDetectedFields(await getDetectedFields(templateId));
  }, [templateId]);

  const handleCreateManualField = useCallback(
    async (
      field: Parameters<typeof addTemplateField>[1],
      _target: { kind: 'group'; id: string } | { kind: 'table'; id: string } | null
    ) => {
      if (!templateId) return null;
      return addTemplateField(templateId, field);
    },
    [templateId]
  );

  const handleUpdateField = useCallback(
    async (fieldId: string, patch: { labelCandidate?: string; type?: string }) => {
      if (!templateId) return;
      await updateTemplateField(templateId, fieldId, patch);
      setDetectedFields((current) =>
        current.map((field) => (field.fieldId === fieldId ? { ...field, ...patch } : field))
      );
    },
    [templateId]
  );

  const handleDeleteField = useCallback(
    async (fieldId: string) => {
      if (!templateId) return;
      await deleteTemplateField(templateId, fieldId);
      await reloadFields();
    },
    [templateId, reloadFields]
  );

  const readOnly = templateStatus === 'archived';

  return (
    <SafeAreaView style={styles.root} edges={['left', 'right']}>
      {editMode ? (
        <SetupWizardStepNav activeStep="assign" onSelectStep={(step) => void handleStepNav(step)} />
      ) : null}
      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>PDF wird vorbereitet…</Text>
        </View>
      ) : null}

      {error ? <Text style={styles.error}>{error}</Text> : null}

      {!loading && setupModel ? (
        <SetupAssignStep
          pdfPath={pdfPath}
          detectedFields={detectedFields}
          mappingFields={mappingFields}
          setupModel={setupModel}
          readOnly={readOnly}
          editMode={editMode}
          onChange={handleChange}
          onComplete={handleComplete}
          onBack={() => void handleBack()}
          onFieldsChanged={() => void reloadFields()}
          onUpdateField={handleUpdateField}
          onDeleteField={handleDeleteField}
          onCreateManualField={handleCreateManualField}
        />
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
  }
});
