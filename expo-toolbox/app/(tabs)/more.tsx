import { useCallback, useState } from 'react';
import { StyleSheet, Text, View } from 'react-native';
import { useFocusEffect, useRouter } from 'expo-router';

import { Card, EmptyState, ListItem, PrimaryButton, Screen, Section } from '../../src/components/mobile';
import { colors, spacing, typography } from '../../src/constants/theme';
import { formatRelativeDate } from '../../src/lib/format';
import { documentRepository } from '../../src/repositories';
import type { Document } from '../../src/types/offline';

export default function MoreScreen() {
  const router = useRouter();
  const [docs, setDocs] = useState<Document[]>([]);

  const load = useCallback(async () => {
    setDocs(await documentRepository.getDocuments());
  }, []);

  useFocusEffect(
    useCallback(() => {
      void load();
    }, [load])
  );

  return (
    <Screen title="Mehr" subtitle="Dokumente · Werkzeuge · Einstellungen">
      <Section title="Dokumente">
        {docs.length === 0 ? (
          <EmptyState
            title="Keine Dokumente"
            description="Lokale Dokument-Metadaten erscheinen hier."
          />
        ) : (
          docs.map((doc) => (
            <ListItem
              key={doc.id}
              title={doc.filename}
              subtitle={doc.mime_type}
              meta={formatRelativeDate(doc.updated_at)}
            />
          ))
        )}
      </Section>

      <Section title="Web-Werkzeuge">
        <Card style={styles.toolCard}>
          <Text style={styles.toolTitle}>SiteReport</Text>
          <Text style={styles.toolCopy}>Protokoll-Workflow im eingebetteten Web-Tool.</Text>
          <PrimaryButton label="SiteReport öffnen" onPress={() => router.push('/sitereport')} />
        </Card>
        <Card style={styles.toolCard}>
          <Text style={styles.toolTitle}>Bautagebuch (Web)</Text>
          <Text style={styles.toolCopy}>PDF-Vorlagen und Erfassung im Web-Tool.</Text>
          <PrimaryButton
            label="Bautagebuch öffnen"
            variant="secondary"
            onPress={() => router.push('/bautagebuch')}
          />
        </Card>
      </Section>

      <Section title="Einstellungen">
        <Card>
          <Text style={styles.settingLabel}>Datenhaltung</Text>
          <Text style={styles.settingValue}>Vollständig offline · SQLite lokal</Text>
          <View style={styles.spacer} />
          <Text style={styles.settingLabel}>Sync / Cloud</Text>
          <Text style={styles.settingValue}>Deaktiviert</Text>
          <View style={styles.spacer} />
          <Text style={styles.settingLabel}>Dark Mode</Text>
          <Text style={styles.settingValue}>Vorbereitet · aktuell Light</Text>
        </Card>
      </Section>
    </Screen>
  );
}

const styles = StyleSheet.create({
  toolCard: {
    marginBottom: spacing.sm,
    gap: spacing.sm
  },
  toolTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  toolCopy: {
    ...typography.caption,
    color: colors.muted
  },
  settingLabel: {
    ...typography.label,
    color: colors.muted
  },
  settingValue: {
    ...typography.body,
    color: colors.ink,
    marginTop: 2
  },
  spacer: {
    height: spacing.sm
  }
});
