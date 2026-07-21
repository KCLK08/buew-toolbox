import { useCallback, useState } from 'react';
import { View } from 'react-native';
import { useFocusEffect, useRouter } from 'expo-router';

import { EmptyState, Fab, ListItem, Screen, StatusBadge } from '../../src/components/mobile';
import { formatRelativeDate, formatStatusLabel, statusTone } from '../../src/lib/format';
import { diaryRepository, photoRepository } from '../../src/repositories';
import type { DiaryEntry } from '../../src/types/offline';

export default function DiaryScreen() {
  const router = useRouter();
  const [items, setItems] = useState<(DiaryEntry & { photoCount: number })[]>([]);
  const [refreshing, setRefreshing] = useState(false);

  const load = useCallback(async () => {
    const entries = await diaryRepository.getDiaryEntries();
    const withPhotos = await Promise.all(
      entries.map(async (entry) => {
        const photos = await photoRepository.getPhotos({
          parent_id: entry.id,
          parent_type: 'diary_entry'
        });
        return { ...entry, photoCount: photos.length };
      })
    );
    setItems(withPhotos);
  }, []);

  useFocusEffect(
    useCallback(() => {
      void load();
    }, [load])
  );

  return (
    <View style={{ flex: 1 }}>
      <Screen
        title="Bautagebuch"
        subtitle={`${items.length} Einträge`}
        refreshing={refreshing}
        onRefresh={() => {
          setRefreshing(true);
          void load().finally(() => setRefreshing(false));
        }}
      >
        {items.length === 0 ? (
          <EmptyState
            title="Keine Einträge"
            description="Erfasse den heutigen Bautag mit wenigen Feldern."
            actionLabel="Eintrag anlegen"
            onAction={() => router.push('/diary/new')}
          />
        ) : (
          items.map((entry) => (
            <ListItem
              key={entry.id}
              title={entry.entry_date || formatRelativeDate(entry.created_at)}
              subtitle={entry.title}
              meta={`Fotos: ${entry.photoCount}`}
              badge={<StatusBadge label={formatStatusLabel(entry.status)} tone={statusTone(entry.status)} />}
              onPress={() => router.push(`/diary/${entry.id}`)}
            />
          ))
        )}
      </Screen>
      <Fab accessibilityLabel="Bautagebuch-Eintrag anlegen" onPress={() => router.push('/diary/new')} />
    </View>
  );
}
