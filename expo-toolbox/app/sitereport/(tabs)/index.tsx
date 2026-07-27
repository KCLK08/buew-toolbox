import { useCallback, useEffect, useMemo, useState } from 'react';
import { Pressable, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { PrimaryActionCard, ProtocolCard } from '../../../src/components/sitereport';
import { Card, EmptyState, Screen, Section, StatCard } from '../../../src/components/mobile';
import { colors, spacing, typography } from '../../../src/constants/theme';
import {
  initSiteReportDatabase,
  listExports,
  listProtocols,
  type SiteReportExport,
  type SiteReportProtocol
} from '../../../src/native/sitereport/db/database';
import { protocolStats } from '../../../src/native/sitereport/services/protocolService';

export default function SiteReportDashboardScreen() {
  const router = useRouter();
  const [loading, setLoading] = useState(true);
  const [protocols, setProtocols] = useState<SiteReportProtocol[]>([]);
  const [exportsList, setExportsList] = useState<SiteReportExport[]>([]);

  const load = useCallback(async () => {
    setLoading(true);
    await initSiteReportDatabase();
    const [nextProtocols, nextExports] = await Promise.all([listProtocols(), listExports()]);
    setProtocols(nextProtocols);
    setExportsList(nextExports);
    setLoading(false);
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const stats = useMemo(() => {
    let photos = 0;
    let entries = 0;
    for (const protocol of protocols) {
      const s = protocolStats(protocol);
      photos += s.photoCount;
      entries += s.entryCount;
    }
    return { protocols: protocols.length, photos, entries, exports: exportsList.length };
  }, [exportsList.length, protocols]);

  const recentProtocols = protocols.slice(0, 3);
  const recentExports = exportsList.slice(0, 3);

  return (
    <Screen
      title="SiteReport"
      subtitle="Baustellen-Protokolle"
      toolboxBack
      reserveTabBarSpace
      scroll
      refreshing={loading}
      onRefresh={load}
    >
      <PrimaryActionCard
        title="Neues Protokoll starten"
        subtitle="In 3 Schritten: Projekt, Format und erste Einträge"
        onPress={() => router.push('/sitereport/new-protocol')}
      />

      <View style={styles.statsRow}>
        <StatCard title="Protokolle" value={String(stats.protocols)} icon="📋" />
        <StatCard title="Einträge" value={String(stats.entries)} icon="✏️" />
        <StatCard title="Fotos" value={String(stats.photos)} icon="📷" />
      </View>

      <Card style={styles.linkCard}>
        <Pressable style={styles.linkRow} onPress={() => router.push('/sitereport/protocols')}>
          <Text style={styles.linkLabel}>Alle Protokolle anzeigen</Text>
          <Text style={styles.linkChevron}>›</Text>
        </Pressable>
        <Pressable style={styles.linkRow} onPress={() => router.push('/sitereport/exports')}>
          <Text style={styles.linkLabel}>Export-Verlauf ({stats.exports})</Text>
          <Text style={styles.linkChevron}>›</Text>
        </Pressable>
      </Card>

      {recentProtocols.length > 0 ? (
        <Section
          title="Zuletzt verwendet"
          action={
            <Pressable onPress={() => router.push('/sitereport/protocols')}>
              <Text style={styles.sectionAction}>Alle</Text>
            </Pressable>
          }
        >
          {recentProtocols.map((protocol) => {
            const s = protocolStats(protocol);
            return (
              <ProtocolCard
                key={protocol.id}
                title={protocol.protocolTitle}
                subtitle={protocol.projectName}
                date={protocol.protocolDate}
                entryCount={s.entryCount}
                photoCount={s.photoCount}
                onPress={() => router.push(`/sitereport/protocol/${protocol.id}`)}
              />
            );
          })}
        </Section>
      ) : (
        <EmptyState
          icon="📋"
          title="Noch keine Protokolle"
          description="Starte dein erstes Baustellen-Protokoll mit dem Button oben."
        />
      )}

      {recentExports.length > 0 ? (
        <Section title="Letzte Exporte">
          {recentExports.map((item) => (
            <ProtocolCard
              key={item.id}
              title={item.protocolTitle || item.projectName}
              subtitle={item.projectName}
              date={item.protocolDate}
              meta={[item.pdfPath ? 'PDF' : null, item.xlsxPath ? 'Excel' : null].filter(Boolean).join(' · ')}
              onPress={() => router.push('/sitereport/exports')}
            />
          ))}
        </Section>
      ) : null}
    </Screen>
  );
}

const styles = StyleSheet.create({
  statsRow: {
    flexDirection: 'row',
    gap: spacing.sm
  },
  linkCard: { gap: 0, paddingVertical: 0 },
  linkRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingVertical: spacing.md,
    minHeight: spacing.touchMin
  },
  linkLabel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  linkChevron: {
    fontSize: 22,
    color: colors.accent
  },
  sectionAction: {
    ...typography.label,
    color: colors.accent
  }
});
