import { useCallback, useEffect, useMemo, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter } from 'expo-router';

import { Screen } from '../../../../src/components/mobile';
import { colors, spacing, typography } from '../../../../src/constants/theme';
import { getDetectedFields } from '../../../../src/native/bautagebuch/db/database';
import { useSetupAutosave } from '../../../../src/native/bautagebuch/hooks/useSetupAutosave';
import {
  ensureWizardInitialized,
  getWizardState,
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

      if (getWizardState(bundle.setupModel).step === 'fields') {
        setLoading(false);
        router.replace(`/bautagebuch/setup/${templateId}/fields`);
        return;
      }

      const initialized = ensureWizardInitialized(bundle.setupModel);
      setTemplateName(bundle.template.templateName);
      setTemplateStatus(bundle.template.status);
      setPdfPath(bundle.template.pdfPath);
      setDetectedFields(fields);
      setSetupModel(initialized);

      if (sortedFields.length === 0) {
        const rebuilt = rebuildSectionsFromWizard(initialized, sortedFields);
        setSetupModel(rebuilt);
        await saveSetupModelToFields(rebuilt);
        setLoading(false);
        router.replace(`/bautagebuch/setup/${templateId}/fields`);
        return;
      }

      if (initialized !== bundle.setupModel) {
        schedule(initialized);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Setup konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, [templateId, schedule, router]);

  async function saveSetupModelToFields(model: Record<string, unknown>) {
    setSetupModel(model);
    schedule(model);
    await flush();
  }

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
      router.replace(`/bautagebuch/setup/${templateId}/fields`);
    })();
  };

  const handleFinishLater = async () => {
    await flush();
    router.back();
  };

  const readOnly = templateStatus === 'archived';

  return (
    <Screen
      title="Schritt 1"
      subtitle={templateName ? `Feldzuordnung · ${templateName}` : 'Feldzuordnung'}
      showBack
      scroll={false}
      scrollableHeader
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
          templateId={String(templateId)}
          templateName={templateName}
          pdfPath={pdfPath}
          detectedFields={detectedFields}
          mappingFields={mappingFields}
          setupModel={setupModel}
          readOnly={readOnly}
          onChange={handleChange}
          onComplete={handleComplete}
          onFinishLater={() => void handleFinishLater()}
          onTemplateRenamed={setTemplateName}
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
