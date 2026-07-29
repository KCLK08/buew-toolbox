import { useCallback, useEffect, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter, type Href } from 'expo-router';
import { SafeAreaView } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../src/constants/theme';
import { SetupStructureIntro } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupStructureIntro';
import { useSetupAutosave } from '../../../../src/native/bautagebuch/hooks/useSetupAutosave';
import {
  ensureWizardInitialized,
  getWizardState,
  markStructureIntroSeen,
  shouldShowAssignIntro,
  shouldShowStructureIntro
} from '../../../../src/native/bautagebuch/lib/setup-mapping';
import { getTemplateBundle } from '../../../../src/native/bautagebuch/services/templateService';

export default function SetupStructureIntroScreen() {
  const router = useRouter();
  const { templateId } = useLocalSearchParams<{ templateId: string }>();
  const [loading, setLoading] = useState(true);
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
      if (!shouldShowStructureIntro(initialized)) {
        setLoading(false);
        router.replace(`/bautagebuch/setup/${templateId}/mapping` as Href);
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
    const next = markStructureIntroSeen(setupModel);
    setSetupModel(next);
    schedule(next);
    void (async () => {
      await flush();
      router.replace(`/bautagebuch/setup/${templateId}/mapping` as Href);
    })();
  };

  const handleBack = async () => {
    await flush();
    router.back();
  };

  return (
    <SafeAreaView style={styles.root} edges={['top', 'left', 'right']}>
      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
        </View>
      ) : null}

      {error ? <Text style={styles.error}>{error}</Text> : null}

      {!loading && setupModel ? (
        <SetupStructureIntro onBack={() => void handleBack()} onStart={handleStart} />
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
