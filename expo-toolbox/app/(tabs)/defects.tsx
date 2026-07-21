import { useCallback, useState } from 'react';
import { View } from 'react-native';
import { useFocusEffect, useRouter } from 'expo-router';

import { EmptyState, Fab, ListItem, Screen, StatusBadge } from '../../src/components/mobile';
import {
  formatRelativeDate,
  formatStatusLabel,
  priorityLabel,
  priorityTone,
  statusTone
} from '../../src/lib/format';
import { defectRepository } from '../../src/repositories';
import type { Defect } from '../../src/types/offline';

export default function DefectsScreen() {
  const router = useRouter();
  const [items, setItems] = useState<Defect[]>([]);
  const [refreshing, setRefreshing] = useState(false);

  const load = useCallback(async () => {
    setItems(await defectRepository.getDefects());
  }, []);

  useFocusEffect(
    useCallback(() => {
      void load();
    }, [load])
  );

  return (
    <View style={{ flex: 1 }}>
      <Screen
        title="Mängel"
        subtitle={`${items.length} erfasst`}
        refreshing={refreshing}
        onRefresh={() => {
          setRefreshing(true);
          void load().finally(() => setRefreshing(false));
        }}
      >
        {items.length === 0 ? (
          <EmptyState
            title="Keine Mängel"
            description="Dokumentiere offene Punkte direkt auf der Baustelle."
            actionLabel="Mangel erfassen"
            onAction={() => router.push('/defect/new')}
          />
        ) : (
          items.map((defect) => (
            <ListItem
              key={defect.id}
              title={defect.title}
              subtitle={defect.description || 'Keine Beschreibung'}
              meta={`Aktualisiert: ${formatRelativeDate(defect.updated_at)}`}
              badge={
                <View style={{ flexDirection: 'row', gap: 6, flexWrap: 'wrap', marginTop: 4 }}>
                  <StatusBadge label={formatStatusLabel(defect.status)} tone={statusTone(defect.status)} />
                  <StatusBadge label={priorityLabel(defect.priority)} tone={priorityTone(defect.priority)} />
                </View>
              }
              onPress={() => router.push(`/defect/${defect.id}`)}
            />
          ))
        )}
      </Screen>
      <Fab accessibilityLabel="Mangel anlegen" onPress={() => router.push('/defect/new')} />
    </View>
  );
}
