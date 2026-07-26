import { useCallback, useEffect, useMemo, useState } from 'react';
import { ActivityIndicator, Alert, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter } from 'expo-router';

import { PrimaryButton, Screen } from '../../../../src/components/mobile';
import { colors, spacing, typography } from '../../../../src/constants/theme';
import { SetupEditor } from '../../../../src/native/bautagebuch/components/SetupEditor';
import { SetupFieldSettingsStep } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupFieldSettingsStep';
import { getDetectedFields, saveSetupModel } from '../../../../src/native/bautagebuch/db/database';
import { useSetupAutosave } from '../../../../src/native/bautagebuch/hooks/useSetupAutosave';
import { hasTableSections, getWizardState } from '../../../../src/native/bautagebuch/lib/setup-mapping';
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
  const [showPreview, setShowPreview] = useState(false);
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
      setDetectedFields(await getDetectedFields(templateId));
      setSetupModel(bundle.setupModel);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Setup konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, [templateId]);

  useEffect(() => {
    void load();
  }, [load]);

  useEffect(() => {
    if (loading || !setupModel || !templateId) return;
    const legacy = templateKind === 'builtin-etb' || hasTableSections(setupModel);
    if (legacy) return;
    if (getWizardState(setupModel).step !== 'fields') {
      router.replace(`/bautagebuch/setup/${templateId}/mapping`);
    }
  }, [loading, setupModel, templateKind, templateId, router]);

  const validationIssues = useMemo(
    () => (setupModel ? validateSetupModel(setupModel) : []),
    [setupModel]
  );

  const useLegacyEditor = Boolean(
    setupModel && (templateKind === 'builtin-etb' || hasTableSections(setupModel))
  );
  const readOnly = templateStatus === 'archived';

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
      const readyModel = {
        ...setupModel,
        status: 'ready',
        updatedAt: new Date().toISOString()
      };
      await saveSetupModel(templateId, readyModel, 'ready');
      Alert.alert('Setup abgeschlossen', 'Die Vorlage ist jetzt startbereit und kann aktiviert werden.');
      router.replace('/bautagebuch/setup');
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Setup konnte nicht abgeschlossen werden.');
    } finally {
      setSaving(false);
    }
  };

  const footer =
    !loading && setupModel && !readOnly ? (
      <View style={styles.footer}>
        {pdfPath ? (
          <PrimaryButton
            label={showPreview ? 'Vorschau ausblenden' : 'Vorschau einblenden'}
            variant="secondary"
            onPress={() => setShowPreview((value) => !value)}
          />
        ) : null}
        <PrimaryButton
          label={saving ? 'Wird abgeschlossen…' : 'Setup abschließen'}
          disabled={saving || validationIssues.length > 0}
          onPress={() => void handleFinish()}
        />
      </View>
    ) : pdfPath && !loading && setupModel ? (
      <PrimaryButton
        label={showPreview ? 'Vorschau ausblenden' : 'Vorschau einblenden'}
        variant="secondary"
        onPress={() => setShowPreview((value) => !value)}
      />
    ) : undefined;

  return (
    <Screen
      title="Schritt 2"
      subtitle={templateName ? `Feldeinstellungen · ${templateName}` : 'Feldeinstellungen'}
      showBack
      scroll={false}
      contentStyle={styles.screenContent}
      footer={footer}
    >
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
          showPreview={showPreview}
          onClosePreview={() => setShowPreview(false)}
          onSelectEdit={() => undefined}
          onSetActive={() => undefined}
          onImport={() => undefined}
        />
      ) : null}

      {!loading && setupModel && !useLegacyEditor ? (
        <SetupFieldSettingsStep
          pdfPath={pdfPath}
          detectedFields={detectedFields}
          setupModel={setupModel}
          validationIssues={validationIssues}
          readOnly={readOnly}
          showPreview={showPreview}
          onChange={handleChange}
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
  },
  footer: {
    gap: spacing.sm
  }
});
