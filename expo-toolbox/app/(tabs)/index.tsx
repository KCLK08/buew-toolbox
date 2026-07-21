import { useCallback, useState } from 'react';
import { StyleSheet, Text, View } from 'react-native';
import { useFocusEffect, useRouter } from 'expo-router';

import { OfflineStatusBanner } from '../../src/components/OfflineStatusBanner';
import { Screen, Section, StatCard } from '../../src/components/mobile';
import { colors, spacing, typography } from '../../src/constants/theme';
import { useOfflineBootstrap } from '../../src/hooks/useOfflineBootstrap';
import { formatRelativeDate } from '../../src/lib/format';
import {
  defectRepository,
  diaryRepository,
  projectRepository
} from '../../src/repositories';

export default function DashboardScreen() {
  const router = useRouter();
  const offline = useOfflineBootstrap();
  const [bannerDismissed, setBannerDismissed] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const [projectCount, setProjectCount] = useState(0);
  const [openDefects, setOpenDefects] = useState(0);
  const [todayDiary, setTodayDiary] = useState('—');

  const load = useCallback(async () => {
    const [projects, defects, diaries] = await Promise.all([
      projectRepository.getProjects(),
      defectRepository.getDefects(),
      diaryRepository.getDiaryEntries()
    ]);
    setProjectCount(projects.filter((p) => p.status === 'active' || p.status === 'draft').length);
    setOpenDefects(defects.filter((d) => d.status !== 'completed' && d.status !== 'archived').length);

    const today = new Date().toISOString().slice(0, 10);
    const todays = diaries.find(
      (entry) => entry.entry_date === today || entry.created_at.startsWith(today)
    );
    setTodayDiary(todays ? todays.title : 'Kein Eintrag');
  }, []);

  useFocusEffect(
    useCallback(() => {
      void load();
    }, [load])
  );

  const onRefresh = () => {
    setRefreshing(true);
    void load().finally(() => setRefreshing(false));
  };

  return (
    <Screen
      title="BÜW-Toolbox"
      subtitle="Offline · lokal auf diesem Gerät"
      contentStyle={styles.content}
      refreshing={refreshing}
      onRefresh={onRefresh}
    >
      {!bannerDismissed ? (
        <OfflineStatusBanner
          report={offline.report}
          error={offline.error}
          onDismiss={() => setBannerDismissed(true)}
        />
      ) : null}

      <View>
        <Text style={styles.greeting}>Baustellen-Übersicht</Text>
        <Text style={styles.hint}>Stand: {formatRelativeDate(new Date().toISOString())}</Text>
      </View>

      <Section title="Heute">
        <View style={styles.stats}>
          <StatCard
            title="Projekte"
            value={String(projectCount)}
            icon="▤"
            tone="accent"
            onPress={() => router.push('/(tabs)/projects')}
          />
          <StatCard
            title="Offene Mängel"
            value={String(openDefects)}
            icon="!"
            tone={openDefects > 0 ? 'warning' : 'default'}
            onPress={() => router.push('/(tabs)/defects')}
          />
        </View>
        <View style={styles.statsSingle}>
          <StatCard
            title="Heute Bautag"
            value={todayDiary}
            icon="✎"
            onPress={() => router.push('/(tabs)/diary')}
          />
        </View>
      </Section>

      <Section title="Schnellzugriff">
        <View style={styles.stats}>
          <StatCard title="Neues Projekt" value="+" onPress={() => router.push('/project/new')} />
          <StatCard title="Neuer Eintrag" value="+" onPress={() => router.push('/diary/new')} />
        </View>
      </Section>
    </Screen>
  );
}

const styles = StyleSheet.create({
  content: {
    gap: spacing.md
  },
  greeting: {
    ...typography.title,
    color: colors.ink
  },
  hint: {
    ...typography.caption,
    color: colors.muted,
    marginBottom: spacing.xs
  },
  stats: {
    flexDirection: 'row',
    gap: spacing.sm
  },
  statsSingle: {
    marginTop: spacing.sm
  }
});
