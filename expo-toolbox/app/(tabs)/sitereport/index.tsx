import { useCallback, useEffect, useMemo, useState } from 'react';
import { Pressable, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { PrimaryActionCard, ProtocolCard, SectionHeader } from '../../../src/components/sitereport';
import { EmptyState, Screen, StatCard } from '../../../src/components/mobile';
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
    <Screen title="SiteReport" subtitle="Baustellen-Protokolle" scroll refreshing={loading} onRefresh={load}>
      <View style={styles.statsRow}>
        <StatCard title="Protokolle" value={String(stats.protocols)} icon="📋" />
        <StatCard title="Einträge" value={String(stats.entries)} icon="✏️" />
      </View>
      <View style={styles.statsRow}>
        <StatCard title="Fotos" value={String(stats.photos)} icon="📷" />
        <StatCard
          title="Exporte"
          value={String(stats.exports)}
          icon="📤"
          tone="accent"
          onPress={() => router.push('/sitereport/exports')}
        />
      </View>

      <PrimaryActionCard
        title="Neues Protokoll starten"
        subtitle="Geführter Setup in 3 Schritten"
        onPress={() => router.push('/sitereport/new-protocol')}
      />

      <Pressable style={styles.linkRow} onPress={() => router.push('/sitereport/protocols')}>
        <Text style={styles.linkLabel}>Alle Protokolle anzeigen</Text>
        <Text style={styles.linkChevron}>›</Text>
      </Pressable>

      {recentProtocols.length > 0 ? (
        <View style={styles.section}>
          <SectionHeader
            title="Zuletzt verwendet"
            actionLabel="Alle"
            onAction={() => router.push('/sitereport/protocols')}
          />
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
        </View>
      ) : (
        <EmptyState
          icon="📋"
          title="Noch keine Protokolle"
          description="Starte dein erstes Baustellen-Protokoll mit dem Button oben."
        />
      )}

      {recentExports.length > 0 ? (
        <View style={styles.section}>
          <SectionHeader
            title="Letzte Exporte"
            actionLabel="Alle"
            onAction={() => router.push('/sitereport/exports')}
          />
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
        </View>
      ) : null}
    </Screen>
  );
}

const styles = StyleSheet.create({
  statsRow: {
    flexDirection: 'row',
    gap: spacing.sm,
    marginBottom: spacing.sm
  },
  linkRow: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingVertical: spacing.sm,
    marginTop: spacing.sm,
    marginBottom: spacing.md
  },
  linkLabel: {
    ...typography.bodyStrong,
    color: colors.accent
  },
  linkChevron: {
    fontSize: 22,
    color: colors.accent
  },
  section: {
    marginBottom: spacing.lg
  }
});
