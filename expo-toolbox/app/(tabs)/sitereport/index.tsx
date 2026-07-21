import { useCallback, useEffect, useState } from 'react';
import { StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { EmptyState, Fab, ListItem, PrimaryButton, Screen, TextField } from '../../../src/components/mobile';
import { colors, typography } from '../../../src/constants/theme';
import {
  createProtocol,
  initSiteReportDatabase,
  listProtocols,
  todayDe
} from '../../../src/native/sitereport/db/database';

export default function SiteReportHomeScreen() {
  const router = useRouter();
  const [loading, setLoading] = useState(true);
  const [protocols, setProtocols] = useState<Awaited<ReturnType<typeof listProtocols>>>([]);
  const [title, setTitle] = useState('');
  const [project, setProject] = useState('');

  const load = useCallback(async () => {
    setLoading(true);
    await initSiteReportDatabase();
    setProtocols(await listProtocols());
    setLoading(false);
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const startProtocol = async () => {
    const protocol = await createProtocol({
      protocolTitle: title.trim() || 'Neues Protokoll',
      projectName: project.trim() || 'Projekt',
      protocolDate: todayDe()
    });
    setTitle('');
    setProject('');
    router.push(`/sitereport/protocol/${protocol.id}`);
  };

  return (
    <Screen title="SiteReport" subtitle="Foto-Protokolle mit Export" scroll refreshing={loading} onRefresh={load}>
      <View style={styles.newCard}>
        <Text style={styles.cardTitle}>Neues Protokoll</Text>
        <TextField label="Titel" value={title} onChangeText={setTitle} />
        <TextField label="Projekt" value={project} onChangeText={setProject} />
        <PrimaryButton label="Protokoll starten" onPress={() => void startProtocol()} />
      </View>

      {protocols.length === 0 ? (
        <EmptyState title="Keine Protokolle" description="Erstelle ein neues Foto-Protokoll für die Baustelle." />
      ) : (
        protocols.map((protocol) => (
          <ListItem
            key={protocol.id}
            title={protocol.protocolTitle}
            subtitle={protocol.projectName}
            meta={`${protocol.protocolDate} · ${protocol.entries.length} Einträge`}
            onPress={() => router.push(`/sitereport/protocol/${protocol.id}`)}
          />
        ))
      )}

      <Fab label="+" onPress={() => void startProtocol()} accessibilityLabel="Neues Protokoll" />
    </Screen>
  );
}

const styles = StyleSheet.create({
  newCard: { gap: 12, marginBottom: 8 },
  cardTitle: { ...typography.bodyStrong, color: colors.ink }
});
