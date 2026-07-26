import { useCallback, useEffect, useState } from 'react';
import { ActivityIndicator, Alert, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { Screen } from '../../../src/components/mobile';
import { colors, spacing, typography } from '../../../src/constants/theme';
import { TemplateOverviewList } from '../../../src/native/bautagebuch/components/setup-wizard/TemplateOverviewList';
import { getSetupModel } from '../../../src/native/bautagebuch/db/database';
import { resolveSetupEntryPath } from '../../../src/native/bautagebuch/lib/setup-mapping';
import {
  archiveTemplate,
  ensureBuiltinTemplate,
  importTemplateFromDocument,
  listManagedTemplates,
  resolveActiveTemplateId,
  setActiveTemplateId
} from '../../../src/native/bautagebuch/services/templateService';

export default function BautagebuchConfigTabScreen() {
  const router = useRouter();
  const [loading, setLoading] = useState(true);
  const [templates, setTemplates] = useState<Awaited<ReturnType<typeof listManagedTemplates>>>([]);
  const [activeTemplateId, setActiveTemplateIdState] = useState('');
  const [importing, setImporting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      await ensureBuiltinTemplate();
      const [templateList, activeId] = await Promise.all([
        listManagedTemplates(),
        resolveActiveTemplateId()
      ]);
      setTemplates(templateList);
      setActiveTemplateIdState(activeId);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Vorlagen konnten nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const navigateToTemplate = async (templateId: string) => {
    const setupModel = await getSetupModel(templateId);
    const template = templates.find((entry) => entry.templateId === templateId);
    if (!setupModel || !template) {
      Alert.alert('Vorlage', 'Setup-Modell nicht gefunden.');
      return;
    }
    router.push(resolveSetupEntryPath(templateId, setupModel, template.templateKind));
  };

  const handleImport = async () => {
    setImporting(true);
    setError(null);
    try {
      const result = await importTemplateFromDocument();
      if (!result) return;
      const templateList = await listManagedTemplates();
      setTemplates(templateList);
      const setupModel = await getSetupModel(result.templateId);
      if (setupModel) {
        router.push(resolveSetupEntryPath(result.templateId, setupModel));
      }
    } catch (err) {
      Alert.alert('Import', err instanceof Error ? err.message : 'Vorlage konnte nicht importiert werden.');
    } finally {
      setImporting(false);
    }
  };

  const handleActivate = async (templateId: string) => {
    setError(null);
    try {
      await setActiveTemplateId(templateId);
      setActiveTemplateIdState(templateId);
      Alert.alert('Aktiv', 'Diese Vorlage wird jetzt für neue Bautagebücher verwendet.');
      await load();
    } catch (err) {
      Alert.alert('Aktivieren', err instanceof Error ? err.message : 'Vorlage konnte nicht aktiviert werden.');
    }
  };

  const handleArchive = async (templateId: string) => {
    Alert.alert('Archivieren', 'Vorlage wirklich archivieren?', [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Archivieren',
        style: 'destructive',
        onPress: () => {
          void archiveTemplate(templateId)
            .then(() => load())
            .catch((err) => {
              Alert.alert('Archiv', err instanceof Error ? err.message : 'Archivierung fehlgeschlagen.');
            });
        }
      }
    ]);
  };

  return (
    <Screen
      title="Vorlagen-Setup"
      subtitle="PDF-Vorlagen verwalten und konfigurieren"
      toolboxBack
      reserveTabBarSpace
      refreshing={loading}
      onRefresh={load}
    >
      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Vorlagen werden geladen…</Text>
        </View>
      ) : null}

      {error ? <Text style={styles.error}>{error}</Text> : null}

      {!loading ? (
        <TemplateOverviewList
          templates={templates}
          activeTemplateId={activeTemplateId}
          importing={importing}
          onImport={() => void handleImport()}
          onOpen={(templateId) => void navigateToTemplate(templateId)}
          onContinueSetup={(templateId) => void navigateToTemplate(templateId)}
          onActivate={(templateId) => void handleActivate(templateId)}
          onArchive={(templateId) => void handleArchive(templateId)}
        />
      ) : null}
    </Screen>
  );
}

const styles = StyleSheet.create({
  center: {
    alignItems: 'center',
    justifyContent: 'center',
    gap: spacing.sm,
    padding: spacing.xl
  },
  muted: {
    ...typography.caption,
    color: colors.muted
  },
  error: {
    ...typography.caption,
    color: colors.danger,
    paddingHorizontal: spacing.pageX,
    paddingBottom: spacing.sm
  }
});
