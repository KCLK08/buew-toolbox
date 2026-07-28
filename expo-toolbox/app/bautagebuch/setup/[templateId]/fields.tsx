import { useCallback, useEffect, useMemo, useState } from 'react';
import { ActivityIndicator, Alert, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter } from 'expo-router';

import { PrimaryButton, Screen } from '../../../../src/components/mobile';
import { colors, spacing, typography } from '../../../../src/constants/theme';
import { PreviewOverlayPanel } from '../../../../src/native/bautagebuch/components/PreviewOverlayPanel';
import { SetupEditor } from '../../../../src/native/bautagebuch/components/SetupEditor';
import { SetupPdfFieldPreview } from '../../../../src/native/bautagebuch/components/SetupPdfFieldPreview';
import { SetupFieldSettingsStep } from '../../../../src/native/bautagebuch/components/setup-wizard/SetupFieldSettingsStep';
import { getDetectedFields, saveSetupModel } from '../../../../src/native/bautagebuch/db/database';
import { useSetupAutosave } from '../../../../src/native/bautagebuch/hooks/useSetupAutosave';
import {
  hasTableSections,
  getWizardState,
  listSetupSections,
  resolveSetupEntryPath
} from '../../../../src/native/bautagebuch/lib/setup-mapping';
import { validateSetupModel } from '../../../../src/native/bautagebuch/lib/setup-model.js';
import {
  getTemplateBundle,
  setActiveTemplateId
} from '../../../../src/native/bautagebuch/services/templateService';

export type SetupPreviewField = {
  fieldId: string;
  label: string;
  page: number;
};

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
  const [previewField, setPreviewField] = useState<SetupPreviewField | null>(null);
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
      router.replace(resolveSetupEntryPath(String(templateId), setupModel, templateKind));
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

  const showFinishDialog = (readyModel: Record<string, unknown>) => {
    const sections = listSetupSections(readyModel);
    const fieldCount = sections.reduce((sum, section) => sum + section.fields.length, 0);
    const groupCount = sections.length;

    Alert.alert(
      'Vorlage fertig eingerichtet',
      `Die Vorlage kann jetzt für neue Bautagebücher verwendet werden.\n\nFelder: ${fieldCount}\nGruppen: ${groupCount}`,
      [
        {
          text: 'Später aktivieren',
          style: 'cancel',
          onPress: () => router.replace('/bautagebuch/config')
        },
        {
          text: 'Vorlage aktivieren',
          onPress: () => {
            void (async () => {
              try {
                if (templateId) {
                  await setActiveTemplateId(templateId);
                }
                router.replace('/bautagebuch/config');
              } catch (err) {
                Alert.alert(
                  'Aktivieren',
                  err instanceof Error ? err.message : 'Vorlage konnte nicht aktiviert werden.'
                );
                router.replace('/bautagebuch/config');
              }
            })();
          }
        }
      ]
    );
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
      setSetupModel(readyModel);
      showFinishDialog(readyModel);
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
            compact
            label={showPreview ? 'Vorschau aus' : 'Vorschau ein'}
            variant="ghost"
            onPress={() => setShowPreview((value) => !value)}
            style={styles.footerSide}
          />
        ) : (
          <View style={styles.footerSide} />
        )}
        <PrimaryButton
          compact
          label={saving ? 'Speichern…' : 'Vorlage speichern'}
          disabled={saving || validationIssues.length > 0}
          onPress={() => void handleFinish()}
          style={styles.footerPrimary}
        />
      </View>
    ) : pdfPath && !loading && setupModel ? (
      <PrimaryButton
        compact
        label={showPreview ? 'Vorschau aus' : 'Vorschau ein'}
        variant="secondary"
        onPress={() => setShowPreview((value) => !value)}
      />
    ) : undefined;

  return (
    <Screen
      title="Schritt 3 von 3"
      subtitle={templateName ? `Feldeinstellungen · ${templateName}` : 'Feldeinstellungen'}
      showBack
      scroll={false}
      scrollableHeader
      compactFooter
      contentStyle={styles.screenContent}
      overlay={
        showPreview && pdfPath ? (
          <PreviewOverlayPanel title="Live-PDF-Vorschau" onClose={() => setShowPreview(false)}>
            <SetupPdfFieldPreview
              variant="overlay"
              emphasizeActiveHighlight
              pdfPath={pdfPath}
              detectedFields={detectedFields}
              activeFieldId={previewField?.fieldId || null}
              activeFieldLabel={previewField?.label || null}
              activeFieldPage={previewField?.page || 1}
            />
          </PreviewOverlayPanel>
        ) : null
      }
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
          onActiveFieldChange={setPreviewField}
          onTemplateRenamed={setTemplateName}
          onSelectEdit={() => undefined}
          onSetActive={() => undefined}
          onImport={() => undefined}
        />
      ) : null}

      {!loading && setupModel && !useLegacyEditor ? (
        <SetupFieldSettingsStep
          templateId={String(templateId)}
          templateName={templateName}
          pdfPath={pdfPath}
          detectedFields={detectedFields}
          setupModel={setupModel}
          validationIssues={validationIssues}
          readOnly={readOnly}
          showPreview={showPreview}
          onActiveFieldChange={setPreviewField}
          onChange={handleChange}
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
  },
  footer: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: spacing.sm
  },
  footerSide: {
    flex: 1
  },
  footerPrimary: {
    flex: 1.4
  }
});
