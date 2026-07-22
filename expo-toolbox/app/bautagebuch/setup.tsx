import { useCallback, useEffect, useMemo, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { PrimaryButton, Screen } from '../../src/components/mobile';
import { colors, spacing, typography } from '../../src/constants/theme';
import { SetupEditor } from '../../src/native/bautagebuch/components/SetupEditor';
import { getDetectedFields, saveSetupModel } from '../../src/native/bautagebuch/db/database';
import { useSetupAutosave } from '../../src/native/bautagebuch/hooks/useSetupAutosave';
import { getActiveTemplateBundle } from '../../src/native/bautagebuch/services/templateService';
import { validateSetupModel } from '../../src/native/bautagebuch/lib/setup-model.js';
import type { DetectedField } from '../../src/native/bautagebuch/types';

export default function BautagebuchSetupScreen() {
  const router = useRouter();
  const [loading, setLoading] = useState(true);
  const [templateId, setTemplateId] = useState('');
  const [templateName, setTemplateName] = useState('');
  const [templatePdfPath, setTemplatePdfPath] = useState<string | null>(null);
  const [detectedFields, setDetectedFields] = useState<DetectedField[]>([]);
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const { schedule, flush } = useSetupAutosave(templateId);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const bundle = await getActiveTemplateBundle();
      setTemplateId(bundle.template.templateId);
      setTemplateName(bundle.template.templateName);
      setTemplatePdfPath(bundle.template.pdfPath);
      setSetupModel(bundle.setupModel);
      setDetectedFields(await getDetectedFields(bundle.template.templateId));
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Setup konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const handleChange = (next: Record<string, unknown>) => {
    setSetupModel(next);
    setInfo('Änderungen werden gespeichert…');
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
      const readyModel = {
        ...setupModel,
        status: 'ready',
        updatedAt: new Date().toISOString()
      };
      await saveSetupModel(templateId, readyModel, 'ready');
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
      subtitle={templateName || 'eBTB-Vorlage anpassen'}
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
