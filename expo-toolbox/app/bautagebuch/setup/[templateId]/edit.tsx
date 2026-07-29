import { useCallback, useEffect, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter, type Href } from 'expo-router';
import { SafeAreaView } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../src/constants/theme';
import { TemplateEditSelection } from '../../../../src/native/bautagebuch/components/setup-wizard/TemplateEditSelection';
import { useSetupAutosave } from '../../../../src/native/bautagebuch/hooks/useSetupAutosave';
import {
  enterEditMode,
  resolveSetupEditStepPath,
  resolveTemplateDetailPath
} from '../../../../src/native/bautagebuch/lib/setup-mapping';
import type { SetupWizardStep } from '../../../../src/native/bautagebuch/types';
import { getTemplateBundle } from '../../../../src/native/bautagebuch/services/templateService';

export default function TemplateEditScreen() {
  const router = useRouter();
  const { templateId } = useLocalSearchParams<{ templateId: string }>();
  const [loading, setLoading] = useState(true);
  const [templateName, setTemplateName] = useState('');
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [error, setError] = useState<string | null>(null);
  const { schedule, flush } = useSetupAutosave(String(templateId || ''));

  const load = useCallback(async () => {
    if (!templateId) return;
    setLoading(true);
    setError(null);
    try {
      const bundle = await getTemplateBundle(templateId);
      setTemplateName(bundle.template.templateName);
      setSetupModel(bundle.setupModel);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Vorlage konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, [templateId]);

  useEffect(() => {
    void load();
  }, [load]);

  const handleSelectStep = async (step: SetupWizardStep) => {
    if (!templateId || !setupModel) return;
    const next = enterEditMode(setupModel, step);
    setSetupModel(next);
    schedule(next);
    await flush();
    router.push(resolveSetupEditStepPath(String(templateId), step));
  };

  return (
    <SafeAreaView style={styles.root} edges={['left', 'right', 'bottom']}>
      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Bearbeitung wird vorbereitet…</Text>
        </View>
      ) : null}

      {error ? <Text style={styles.error}>{error}</Text> : null}

      {!loading && setupModel ? (
        <TemplateEditSelection
          templateName={templateName}
          onBack={() => router.replace(resolveTemplateDetailPath(String(templateId)) as Href)}
          onSelectStep={(step) => void handleSelectStep(step)}
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
