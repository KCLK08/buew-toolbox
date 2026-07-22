import { useCallback, useEffect, useMemo, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter } from 'expo-router';

import { Screen } from '../../../../src/components/mobile';
import { colors, spacing, typography } from '../../../../src/constants/theme';
import { getDetectedFields } from '../../../../src/native/bautagebuch/db/database';
import { useSetupAutosave } from '../../../../src/native/bautagebuch/hooks/useSetupAutosave';
import {
  ensureWizardInitialized,
  rebuildSectionsFromWizard,
  sortMappingFields
} from '../../../../src/native/bautagebuch/lib/setup-mapping';
import { SetupMappingStep } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupMappingStep';
import { getTemplateBundle } from '../../../../src/native/bautagebuch/services/templateService';

export default function SetupMappingScreen() {
  const router = useRouter();
  const { templateId } = useLocalSearchParams<{ templateId: string }>();
  const [loading, setLoading] = useState(true);
  const [templateName, setTemplateName] = useState('');
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
      const initialized = ensureWizardInitialized(bundle.setupModel);
      setTemplateName(bundle.template.templateName);
      setPdfPath(bundle.template.pdfPath);
      setDetectedFields(fields);
      setSetupModel(initialized);
      schedule(initialized);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Setup konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, [templateId, schedule]);

  useEffect(() => {
    void load();
  }, [load]);

  const mappingFields = useMemo(() => sortMappingFields(detectedFields), [detectedFields]);

  const handleChange = (next: Record<string, unknown>) => {
    setSetupModel(next);
    schedule(next);
  };

  const goToFields = async (model: Record<string, unknown>) => {
    await flush();
    router.replace(`/bautagebuch/setup/${templateId}/fields`);
  };

  const handleComplete = () => {
    if (!setupModel) return;
    const rebuilt = rebuildSectionsFromWizard(setupModel, mappingFields);
    setSetupModel(rebuilt);
    schedule(rebuilt);
    void goToFields(rebuilt);
  };

  const handleFinishLater = async () => {
    await flush();
    router.back();
  };

  return (
    <Screen
      title="Schritt 1"
      subtitle={templateName ? `Feldzuordnung · ${templateName}` : 'Feldzuordnung'}
      showBack
      scroll={false}
      contentStyle={styles.screenContent}
    >
      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>PDF wird vorbereitet…</Text>
        </View>
      ) : null}

      {error ? <Text style={styles.error}>{error}</Text> : null}

      {!loading && setupModel ? (
        <SetupMappingStep
          pdfPath={pdfPath}
          detectedFields={detectedFields}
          mappingFields={mappingFields}
          setupModel={setupModel}
          onChange={handleChange}
          onComplete={handleComplete}
          onFinishLater={() => void handleFinishLater()}
        />
      ) : null}
    </Screen>
  );
}

const styles = StyleSheet.create({
  screenContent: {
    flex: 1,
    paddingHorizontal: 0,
    paddingTop: 0
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
