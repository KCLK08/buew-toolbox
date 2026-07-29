import { useCallback, useEffect, useState } from 'react';
import { ActivityIndicator, StyleSheet, Text, View } from 'react-native';
import { useLocalSearchParams, useRouter, type Href } from 'expo-router';
import { SafeAreaView } from 'react-native-safe-area-context';

import { colors, spacing, typography } from '../../../../src/constants/theme';
import { TemplateDetailView } from '../../../../src/native/bautagebuch/components/setup-wizard/TemplateDetailView';
import { getDetectedFields } from '../../../../src/native/bautagebuch/db/database';
import { resolveTemplateEditPath } from '../../../../src/native/bautagebuch/lib/setup-mapping';
import { getTemplateBundle } from '../../../../src/native/bautagebuch/services/templateService';

export default function TemplateDetailScreen() {
  const router = useRouter();
  const { templateId } = useLocalSearchParams<{ templateId: string }>();
  const [loading, setLoading] = useState(true);
  const [templateName, setTemplateName] = useState('');
  const [templateStatus, setTemplateStatus] = useState('');
  const [setupModel, setSetupModel] = useState<Record<string, unknown> | null>(null);
  const [detectedFields, setDetectedFields] = useState<Awaited<ReturnType<typeof getDetectedFields>>>([]);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    if (!templateId) return;
    setLoading(true);
    setError(null);
    try {
      const bundle = await getTemplateBundle(templateId);
      setTemplateName(bundle.template.templateName);
      setTemplateStatus(bundle.template.status);
      setSetupModel(bundle.setupModel);
      setDetectedFields(await getDetectedFields(templateId));
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Vorlage konnte nicht geladen werden.');
    } finally {
      setLoading(false);
    }
  }, [templateId]);

  useEffect(() => {
    void load();
  }, [load]);

  const readOnly = templateStatus === 'archived';

  return (
    <SafeAreaView style={styles.root} edges={['left', 'right', 'bottom']}>
      {loading ? (
        <View style={styles.center}>
          <ActivityIndicator color={colors.accent} />
          <Text style={styles.muted}>Vorlage wird geladen…</Text>
        </View>
      ) : null}

      {error ? <Text style={styles.error}>{error}</Text> : null}

      {!loading && setupModel ? (
        <TemplateDetailView
          templateName={templateName}
          setupModel={setupModel}
          detectedFields={detectedFields}
          readOnly={readOnly}
          onBack={() => router.replace('/bautagebuch/config' as Href)}
          onEdit={
            readOnly
              ? undefined
              : () => router.push(resolveTemplateEditPath(String(templateId)))
          }
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
