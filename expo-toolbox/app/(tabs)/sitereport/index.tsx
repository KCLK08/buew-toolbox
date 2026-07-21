import { useCallback, useEffect, useMemo, useState } from 'react';
import { StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { DashboardActionCard } from '../../../src/components/sitereport';
import { Fab, ListItem, Screen, StatCard } from '../../../src/components/mobile';
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
    return { protocols: protocols.length, photos, entries };
  }, [protocols]);

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
        <StatCard title="Exporte" value={String(exportsList.length)} icon="📤" tone="accent" onPress={() => router.push('/sitereport/exports')} />
      </View>

      <View style={styles.actionsRow}>
        <DashboardActionCard
          title="Neues Protokoll"
          subtitle="Geführter Setup-Flow"
          icon="➕"
          accent
          onPress={() => router.push('/sitereport/new-protocol')}
        />
        <DashboardActionCard
          title="Protokolle"
          subtitle={`${protocols.length} gespeichert`}
          icon="📁"
          onPress={() => router.push('/sitereport/protocols')}
        />
      </View>

      {recentProtocols.length > 0 ? (
        <View style={styles.section}>
          <View style={styles.sectionHeader}>
            <Text style={styles.sectionTitle}>Aktive Protokolle</Text>
            <Text style={styles.sectionLink} onPress={() => router.push('/sitereport/protocols')}>
              Alle
            </Text>
          </View>
          {recentProtocols.map((protocol) => {
            const s = protocolStats(protocol);
            return (
              <ListItem
                key={protocol.id}
                title={protocol.protocolTitle}
                subtitle={protocol.projectName}
                meta={`${protocol.protocolDate} · ${s.entryCount} Einträge · ${s.photoCount} Fotos`}
                onPress={() => router.push(`/sitereport/protocol/${protocol.id}`)}
              />
            );
          })}
        </View>
      ) : null}

      {recentExports.length > 0 ? (
        <View style={styles.section}>
          <View style={styles.sectionHeader}>
            <Text style={styles.sectionTitle}>Letzte Exporte</Text>
            <Text style={styles.sectionLink} onPress={() => router.push('/sitereport/exports')}>
              Alle
            </Text>
          </View>
          {recentExports.map((item) => (
            <ListItem
              key={item.id}
              title={item.protocolTitle || item.projectName}
              subtitle={`${item.projectName} · ${item.protocolDate}`}
              meta={[item.pdfPath ? 'PDF' : null, item.xlsxPath ? 'Excel' : null].filter(Boolean).join(' · ')}
              onPress={() => router.push('/sitereport/exports')}
            />
          ))}
        </View>
      ) : null}

      <Fab label="+" onPress={() => router.push('/sitereport/new-protocol')} accessibilityLabel="Neues Protokoll" />
    </Screen>
  );
}

const styles = StyleSheet.create({
  statsRow: {
    flexDirection: 'row',
    gap: spacing.sm,
    marginBottom: spacing.sm
  },
  actionsRow: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.sm,
    marginBottom: spacing.lg
  },
  section: {
    gap: spacing.xs,
    marginBottom: spacing.lg
  },
  sectionHeader: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: spacing.xs
  },
  sectionTitle: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  sectionLink: {
    ...typography.bodyStrong,
    color: colors.accent
  }
});
