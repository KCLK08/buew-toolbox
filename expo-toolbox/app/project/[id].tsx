import { useCallback, useState } from 'react';
import { Alert, Image, StyleSheet, Text, View } from 'react-native';
import { useFocusEffect, useLocalSearchParams, useRouter } from 'expo-router';

import {
  EmptyState,
  PrimaryButton,
  Screen,
  Section,
  StatusBadge,
  TextField
} from '../../src/components/mobile';
import { colors, spacing, typography } from '../../src/constants/theme';
import { formatRelativeDate, formatStatusLabel, statusTone } from '../../src/lib/format';
import { captureAndSavePhoto } from '../../src/lib/photoCapture';
import { resolveDocumentUri } from '../../src/storage/fileService';
import { photoRepository, projectRepository } from '../../src/repositories';
import type { Photo, Project } from '../../src/types/offline';

export default function ProjectDetailScreen() {
  const { id } = useLocalSearchParams<{ id: string }>();
  const router = useRouter();
  const [project, setProject] = useState<Project | null>(null);
  const [photos, setPhotos] = useState<Photo[]>([]);
  const [name, setName] = useState('');
  const [location, setLocation] = useState('');
  const [description, setDescription] = useState('');
  const [saving, setSaving] = useState(false);
  const [photoBusy, setPhotoBusy] = useState(false);

  const load = useCallback(async () => {
    if (!id) return;
    const item = await projectRepository.getProjectById(id);
    setProject(item);
    if (item) {
      setName(item.name);
      setLocation(item.location);
      setDescription(item.description);
    }
    setPhotos(await photoRepository.getPhotos({ parent_id: id, parent_type: 'project' }));
  }, [id]);

  useFocusEffect(
    useCallback(() => {
      void load();
    }, [load])
  );

  const onSave = async () => {
    if (!id || !name.trim()) return;
    setSaving(true);
    try {
      await projectRepository.updateProject({
        id,
        name: name.trim(),
        location: location.trim(),
        description: description.trim()
      });
      await load();
      Alert.alert('Gespeichert', 'Projekt aktualisiert.');
    } finally {
      setSaving(false);
    }
  };

  const onAddPhoto = async () => {
    if (!id) return;
    setPhotoBusy(true);
    try {
      const photo = await captureAndSavePhoto({ parentId: id, parentType: 'project' });
      if (photo) await load();
    } finally {
      setPhotoBusy(false);
    }
  };

  const onDelete = () => {
    if (!id) return;
    Alert.alert('Projekt löschen', 'Projekt wirklich soft-löschen?', [
      { text: 'Abbrechen', style: 'cancel' },
      {
        text: 'Löschen',
        style: 'destructive',
        onPress: () => {
          void projectRepository.softDeleteProject(id).then(() => router.replace('/(tabs)/projects'));
        }
      }
    ]);
  };

  if (!project) {
    return (
      <Screen title="Projekt" showBack>
        <EmptyState title="Projekt nicht gefunden" actionLabel="Zurück" onAction={() => router.back()} />
      </Screen>
    );
  }

  return (
    <Screen
      title={project.name}
      subtitle={formatRelativeDate(project.updated_at)}
      showBack
      footer={
        <>
          <PrimaryButton label="Änderungen speichern" onPress={() => void onSave()} loading={saving} />
          <PrimaryButton label="Löschen" variant="danger" onPress={onDelete} />
        </>
      }
    >
      <StatusBadge label={formatStatusLabel(project.status)} tone={statusTone(project.status)} />
      <TextField label="Name" value={name} onChangeText={setName} />
      <TextField label="Ort" value={location} onChangeText={setLocation} />
      <TextField
        label="Beschreibung"
        value={description}
        onChangeText={setDescription}
        multiline
        style={{ minHeight: 88, textAlignVertical: 'top' }}
      />

      <Section
        title="Fotos"
        action={
          <PrimaryButton
            label="+ Foto"
            variant="secondary"
            onPress={() => void onAddPhoto()}
            loading={photoBusy}
            style={styles.inlineBtn}
          />
        }
      >
        {photos.length === 0 ? (
          <Text style={styles.emptyPhotos}>Noch keine Fotos</Text>
        ) : (
          <View style={styles.gallery}>
            {photos.map((photo) => (
              <Image
                key={photo.id}
                source={{ uri: resolveDocumentUri(photo.local_path) }}
                style={styles.thumb}
                accessibilityLabel={photo.filename}
              />
            ))}
          </View>
        )}
      </Section>
    </Screen>
  );
}

const styles = StyleSheet.create({
  inlineBtn: {
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.sm
  },
  emptyPhotos: {
    ...typography.caption,
    color: colors.muted
  },
  gallery: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.sm
  },
  thumb: {
    width: '47%',
    aspectRatio: 1,
    borderRadius: spacing.inputRadius,
    backgroundColor: colors.border
  }
});
