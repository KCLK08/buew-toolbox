import { useCallback, useEffect, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter, type Href } from 'expo-router';
import { SafeAreaView, useSafeAreaInsets } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../src/constants/theme';
import { SetupFieldSettingsOnboarding } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupFieldSettingsOnboarding';
import { SetupWizardStepNav } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupWizardStepNav';
import { useSetupAutosave } from '../../../../src/native/bautagebuch/hooks/useSetupAutosave';
import {
  ensureWizardInitialized,
  getWizardState,
  markFieldsIntroSeen,
  shouldShowFieldsIntro
} from '../../../../src/native/bautagebuch/lib/setup-mapping';
import { navigateSetupWizardStep } from '../../../../src/native/bautagebuch/lib/setup-wizard-exit';
import { isSetupStepNavigationActive } from '../../../../src/native/bautagebuch/lib/setup-wizard-nav-session';
import { getTemplateBundle } from '../../../../src/native/bautagebuch/services/templateService';
import { systemBottomInset } from '../../../../src/navigation/systemInsets';

export default function SetupFieldsIntroScreen() {
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const { templateId } = useLocalSearchParams<{ templateId: string }>();
  const [loading, setLoading] = useState(true);
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [error, setError] = useState<string | null>(null);
  const { schedule, flush, isPending } = useSetupAutosave(String(templateId || ''));

  const load = useCallback(async () => {
    if (!templateId) return;
    setLoading(true);
    setError(null);
    try {
      const bundle = await getTemplateBundle(templateId);
      const wizard = getWizardState(bundle.setupModel);
      const kind = bundle.template.templateKind;
      const skipRedirects = isSetupStepNavigationActive();

      if (!skipRedirects) {
        if (wizard.step === 'structure') {
          setLoading(false);
          router.replace(`/bautagebuch/setup/${templateId}/mapping` as Href);
          return;
        }
        if (wizard.step === 'assign') {
          setLoading(false);
          router.replace(`/bautagebuch/setup/${templateId}/assign` as Href);
          return;
        }
      }

      const initialized = ensureWizardInitialized(bundle.setupModel);

      if (!skipRedirects && !shouldShowFieldsIntro(initialized, kind)) {
        setLoading(false);
        router.replace(`/bautagebuch/setup/${templateId}/fields` as Href);
        return;
      }

      setSetupModel(initialized);
      if (initialized !== bundle.setupModel) {
        schedule(initialized);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Einführung konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, [templateId, schedule, router]);

  useEffect(() => {
    void load();
  }, [load]);

  const handleStart = () => {
    if (!setupModel || !templateId) return;
    const next = markFieldsIntroSeen(setupModel);
    setSetupModel(next);
    schedule(next);
    void (async () => {
      await flush();
      router.replace(`/bautagebuch/setup/${templateId}/fields` as Href);
    })();
  };

  const handleBack = async () => {
    await flush();
    router.back();
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

  return (
    <SafeAreaView
      style={[styles.root, { paddingBottom: systemBottomInset(insets) }]}
      edges={['left', 'right']}
    >
      {showWizardNav ? (
        <SetupWizardStepNav
          activeStep="fields"
          disabled={!stepNavReady}
          onSelectStep={(step) => void handleStepNav(step)}
        />
      ) : null}

      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
        </View>
      ) : null}

      {error ? <Text style={styles.error}>{error}</Text> : null}

      {!loading && setupModel ? (
        <SetupFieldSettingsOnboarding onBack={() => void handleBack()} onStart={handleStart} />
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
    justifyContent: 'center'
  },
  error: {
    ...typography.caption,
    color: colors.danger,
    padding: spacing.pageX
  }
});
