import { useCallback, useState } from 'react';
import { View } from 'react-native';
import { useFocusEffect, useRouter } from 'expo-router';

import { EmptyState, Fab, ListItem, Screen, StatusBadge } from '../../src/components/mobile';
import { formatRelativeDate, formatStatusLabel, statusTone } from '../../src/lib/format';
import { projectRepository } from '../../src/repositories';
import type { Project } from '../../src/types/offline';

export default function ProjectsScreen() {
  const router = useRouter();
  const [items, setItems] = useState<Project[]>([]);
  const [refreshing, setRefreshing] = useState(false);

  const load = useCallback(async () => {
    setItems(await projectRepository.getProjects());
  }, []);

  useFocusEffect(
    useCallback(() => {
      void load();
    }, [load])
  );

  return (
    <View style={{ flex: 1 }}>
      <Screen
        title="Projekte"
        subtitle={`${items.length} aktiv`}
        refreshing={refreshing}
        onRefresh={() => {
          setRefreshing(true);
          void load().finally(() => setRefreshing(false));
        }}
      >
        {items.length === 0 ? (
          <EmptyState
            title="Noch keine Projekte"
            description="Lege ein Projekt an, um Bautagebuch und Mängel zuzuordnen."
            actionLabel="Projekt anlegen"
            onAction={() => router.push('/project/new')}
          />
        ) : (
          items.map((project) => (
            <ListItem
              key={project.id}
              title={project.name}
              subtitle={project.location || project.description || 'Kein Ort hinterlegt'}
              meta={`Letzte Änderung: ${formatRelativeDate(project.updated_at)}`}
              badge={<StatusBadge label={formatStatusLabel(project.status)} tone={statusTone(project.status)} />}
              onPress={() => router.push(`/project/${project.id}`)}
            />
          ))
        )}
      </Screen>
      <Fab accessibilityLabel="Projekt anlegen" onPress={() => router.push('/project/new')} />
    </View>
  );
}
