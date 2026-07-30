import { useCallback, useEffect, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter, type Href } from 'expo-router';
import { SafeAreaView } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../src/constants/theme';
import { SetupStructureStep } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupStructureStep';
import { SetupWizardStepNav } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupWizardStepNav';
import { useSetupAutosave } from '../../../../src/native/bautagebuch/hooks/useSetupAutosave';
import {
  ensureWizardInitialized,
  getWizardState,
  resolveSetupEditStepPath,
  resolveTemplateEditPath,
  shouldShowAssignIntro,
  shouldShowStructureIntro
} from '../../../../src/native/bautagebuch/lib/setup-mapping';
import { getTemplateBundle } from '../../../../src/native/bautagebuch/services/templateService';

export default function SetupStructureScreen() {
  const router = useRouter();
  const { templateId } = useLocalSearchParams<{ templateId: string }>();
  const [loading, setLoading] = useState(true);
  const [templateStatus, setTemplateStatus] = useState('');
  const [pdfPath, setPdfPath] = useState<string | null>(null);
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [error, setError] = useState<string | null>(null);
  const { schedule, flush } = useSetupAutosave(String(templateId || ''));

  const load = useCallback(async () => {
    if (!templateId) return;
    setLoading(true);
    setError(null);
    try {
      const bundle = await getTemplateBundle(templateId);
      const wizard = getWizardState(bundle.setupModel);

      if (wizard.step === 'fields') {
        setLoading(false);
        router.replace(`/bautagebuch/setup/${templateId}/fields`);
        return;
      }
      if (wizard.step === 'assign') {
        setLoading(false);
        router.replace(
          shouldShowAssignIntro(bundle.setupModel)
            ? (`/bautagebuch/setup/${templateId}/assign-intro` as Href)
            : (`/bautagebuch/setup/${templateId}/assign` as Href)
        );
        return;
      }

      const initialized = ensureWizardInitialized(bundle.setupModel);
      if (shouldShowStructureIntro(initialized)) {
        setLoading(false);
        router.replace(`/bautagebuch/setup/${templateId}/intro` as Href);
        return;
      }

      setTemplateStatus(bundle.template.status);
      setPdfPath(bundle.template.pdfPath);
      setSetupModel(initialized);

      if (initialized !== bundle.setupModel) {
        schedule(initialized);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Setup konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, [templateId, schedule, router]);

  useEffect(() => {
    void load();
  }, [load]);

  const handleChange = (next: Record<string, unknown>) => {
    setSetupModel(next);
    schedule(next);
  };

  const handleComplete = (next: Record<string, unknown>) => {
    setSetupModel(next);
    schedule(next);
    void (async () => {
      await flush();
      router.replace(
        shouldShowAssignIntro(next)
          ? (`/bautagebuch/setup/${templateId}/assign-intro` as Href)
          : (`/bautagebuch/setup/${templateId}/assign` as Href)
      );
    })();
  };

  const handleBack = async () => {
    await flush();
    const wizard = setupModel ? getWizardState(setupModel) : null;
    if (wizard?.editMode) {
      router.replace(resolveTemplateEditPath(String(templateId)));
      return;
    }
    router.back();
  };

  const editMode = setupModel ? getWizardState(setupModel).editMode === true : false;

  const handleStepNav = async (step: 'structure' | 'assign' | 'fields') => {
    await flush();
    router.replace(resolveSetupEditStepPath(String(templateId), step));
  };

  const readOnly = templateStatus === 'archived';

  return (
    <SafeAreaView style={styles.root} edges={['left', 'right']}>
      {editMode ? (
        <SetupWizardStepNav activeStep="structure" onSelectStep={(step) => void handleStepNav(step)} />
      ) : null}
      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Setup wird vorbereitet…</Text>
        </View>
      ) : null}

      {error ? <Text style={styles.error}>{error}</Text> : null}

      {!loading && setupModel ? (
        <SetupStructureStep
          pdfPath={pdfPath}
          setupModel={setupModel}
          readOnly={readOnly}
          editMode={editMode}
          onChange={handleChange}
          onComplete={handleComplete}
          onBack={() => void handleBack()}
          onNavigateToAssign={() => {
            void flush().then(() => {
              router.replace(`/bautagebuch/setup/${templateId}/assign` as Href);
            });
          }}
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
