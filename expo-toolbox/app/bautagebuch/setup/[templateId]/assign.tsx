import { useCallback, useEffect, useMemo, useState } from 'react';
import { ActivityIndicator, KeyboardAvoidingView, Platform, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter, type Href } from 'expo-router';
import { SafeAreaView, useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../src/constants/theme';
import { KeyboardScrollProvider } from '../../../../src/contexts/KeyboardScrollContext';
import { SetupAssignStep } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupAssignStep';
import { SetupWizardStepNav } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupWizardStepNav';
import { addTemplateField, deleteTemplateField, getDetectedFields, updateTemplateField } from '../../../../src/native/bautagebuch/db/database';
import { useSetupAutosave } from '../../../../src/native/bautagebuch/hooks/useSetupAutosave';
import {
  ensureWizardInitialized,
  getWizardState,
  rebuildSectionsFromWizard,
  resolveFieldsSetupPath,
  shouldShowAssignIntro,
  sortMappingFields
} from '../../../../src/native/bautagebuch/lib/setup-mapping';
import { exitSetupWizardToOverview, navigateSetupWizardStep } from '../../../../src/native/bautagebuch/lib/setup-wizard-exit';
import { consumeSetupStepNavigation } from '../../../../src/native/bautagebuch/lib/setup-wizard-nav-session';
import { prepareSetupStepNavigation } from '../../../../src/native/bautagebuch/lib/setup-wizard-navigation';
import { getTemplateBundle } from '../../../../src/native/bautagebuch/services/templateService';
import { systemBottomInset } from '../../../../src/navigation/systemInsets';

export default function SetupAssignScreen() {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const { templateId } = useLocalSearchParams<{ templateId: string }>();
  const [loading, setLoading] = useState(true);
  const [templateStatus, setTemplateStatus] = useState('');
  const [pdfPath, setPdfPath] = useState<string | null>(null);
  const [detectedFields, setDetectedFields] = useState<Awaited<ReturnType<typeof getDetectedFields>>>([]);
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [error, setError] = useState<string | null>(null);
  const { schedule, flush, isPending } = useSetupAutosave(String(templateId || ''));

  const load = useCallback(async () => {
    if (!templateId) return;
    setLoading(true);
    setError(null);
    try {
      const bundle = await getTemplateBundle(templateId);
      const fields = await getDetectedFields(templateId);
      const wizard = getWizardState(bundle.setupModel);
      const steppedInViaNav = consumeSetupStepNavigation('assign');

      if (!steppedInViaNav) {
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
      }

      const initialized = ensureWizardInitialized(bundle.setupModel);
      if (!steppedInViaNav && shouldShowAssignIntro(initialized)) {
        setLoading(false);
        router.replace(`/bautagebuch/setup/${templateId}/assign-intro` as Href);
        return;
      }

      const nextModel = steppedInViaNav
        ? prepareSetupStepNavigation(initialized, 'assign')
        : initialized;

      setTemplateStatus(bundle.template.status);
      setPdfPath(bundle.template.pdfPath);
      setDetectedFields(fields);
      setSetupModel(nextModel);

      if (nextModel !== bundle.setupModel) {
        schedule(nextModel);
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

  const handleBack = () => {
    void exitSetupWizardToOverview({
      autosave: { schedule, flush, isPending },
      router
    });
  };

  const showWizardNav = Boolean(templateId);
  const stepNavReady = Boolean(setupModel && !loading);

  const handleStepNav = async (step: 'structure' | 'assign' | 'fields') => {
    if (!setupModel || !templateId) return;
    await navigateSetupWizardStep({
      templateId: String(templateId),
      step,
      setupModel,
      autosave: { schedule, flush, isPending },
      setSetupModel,
      router
    });
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
    async (
      fieldId: string,
      patch: {
        labelCandidate?: string;
        type?: string;
        geometry?: { page: number; rect: { x: number; y: number; width: number; height: number } } | null;
        page?: number;
      }
    ) => {
      if (!templateId) return;
      const updated = await updateTemplateField(templateId, fieldId, patch);
      if (!updated) return;
      setDetectedFields((current) =>
        current.map((field) => (field.fieldId === fieldId ? updated : field))
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
    <KeyboardScrollProvider footerInset={systemBottomInset(insets)}>
      <KeyboardAvoidingView
        style={styles.flex}
        behavior={Platform.OS === 'ios' ? 'padding' : 'height'}
      >
      <SafeAreaView
        style={[styles.root, { paddingBottom: systemBottomInset(insets) }]}
        edges={['left', 'right']}
      >
      {showWizardNav ? (
        <SetupWizardStepNav
          activeStep="assign"
          disabled={!stepNavReady}
          onSelectStep={(step) => void handleStepNav(step)}
        />
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
          showWizardNav={showWizardNav}
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
      </KeyboardAvoidingView>
    </KeyboardScrollProvider>
  );
}

const styles = StyleSheet.create({
  flex: {
    flex: 1
  },
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
