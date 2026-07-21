import { useCallback, useEffect, useState } from 'react';
import { Alert } from 'react-native';
import { useRouter } from 'expo-router';

import { Screen } from '../../src/components/mobile';
import { SetupEditor } from '../../src/native/bautagebuch/components/SetupEditor';
import { saveSetupModel } from '../../src/native/bautagebuch/db/database';
import { useSetupAutosave } from '../../src/native/bautagebuch/hooks/useSetupAutosave';
import { exportSetupPreviewPdf } from '../../src/native/bautagebuch/services/exportService';
import { getActiveTemplateBundle } from '../../src/native/bautagebuch/services/templateService';
import { validateSetupModel } from '../../src/native/bautagebuch/lib/setup-model.js';

export default function BautagebuchSetupScreen() {
  const router = useRouter();
  const [loading, setLoading] = useState(true);
  const [templateId, setTemplateId] = useState('');
  const [templateName, setTemplateName] = useState('');
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [previewBusy, setPreviewBusy] = useState(false);
  const { schedule, flush } = useSetupAutosave(templateId);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const bundle = await getActiveTemplateBundle();
      setTemplateId(bundle.template.templateId);
      setTemplateName(bundle.template.templateName);
      setSetupModel(bundle.setupModel);
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

  const handlePreview = async () => {
    if (!templateId) return;
    setPreviewBusy(true);
    setError(null);
    try {
      await flush();
      await exportSetupPreviewPdf(templateId);
      setInfo('PDF-Vorschau wurde erstellt und kann geteilt werden.');
    } catch (err) {
      Alert.alert('Vorschau', err instanceof Error ? err.message : 'PDF-Vorschau fehlgeschlagen.');
    } finally {
      setPreviewBusy(false);
    }
  };

  return (
    <Screen
      title="Setup-Editor"
      subtitle="eBTB-Vorlage anpassen"
      showBack
      scroll
      refreshing={loading}
      onRefresh={load}
    >
      {setupModel ? (
        <SetupEditor
          templateName={templateName}
          setupModel={setupModel}
          onChange={handleChange}
          onFinish={() => void handleFinish()}
          onPreview={() => void handlePreview()}
          saving={saving}
          previewBusy={previewBusy}
          info={info}
          error={error}
        />
      ) : null}
    </Screen>
  );
}
