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
import {
  formatRelativeDate,
  formatStatusLabel,
  priorityLabel,
  priorityTone,
  statusTone
} from '../../src/lib/format';
import { captureAndSavePhoto } from '../../src/lib/photoCapture';
import { defectRepository, photoRepository } from '../../src/repositories';
import { resolveDocumentUri } from '../../src/storage/fileService';
import type { Defect, Photo } from '../../src/types/offline';

export default function DefectDetailScreen() {
  const { id } = useLocalSearchParams<{ id: string }>();
  const router = useRouter();
  const [defect, setDefect] = useState<Defect | null>(null);
  const [photos, setPhotos] = useState<Photo[]>([]);
  const [title, setTitle] = useState('');
  const [description, setDescription] = useState('');
  const [saving, setSaving] = useState(false);
  const [photoBusy, setPhotoBusy] = useState(false);

  const load = useCallback(async () => {
    if (!id) return;
    const item = await defectRepository.getDefectById(id);
    setDefect(item);
    if (item) {
      setTitle(item.title);
      setDescription(item.description);
    }
    setPhotos(await photoRepository.getPhotos({ parent_id: id, parent_type: 'defect' }));
  }, [id]);

  useFocusEffect(
    useCallback(() => {
      void load();
    }, [load])
  );

  const onSave = async () => {
    if (!id || !title.trim()) return;
    setSaving(true);
    try {
      await defectRepository.updateDefect({
        id,
        title: title.trim(),
        description: description.trim()
      });
      await load();
      Alert.alert('Gespeichert', 'Mangel aktualisiert.');
    } finally {
      setSaving(false);
    }
  };

  const onAddPhoto = async () => {
    if (!id) return;
    setPhotoBusy(true);
    try {
      const photo = await captureAndSavePhoto({ parentId: id, parentType: 'defect' });
      if (photo) await load();
    } finally {
      setPhotoBusy(false);
    }
  };

  if (!defect) {
    return (
      <Screen title="Mangel" showBack>
        <EmptyState title="Mangel nicht gefunden" actionLabel="Zurück" onAction={() => router.back()} />
      </Screen>
    );
  }

  return (
    <Screen
      title={defect.title}
      subtitle={formatRelativeDate(defect.updated_at)}
      showBack
      footer={
        <>
          <PrimaryButton
            label="+ Foto hinzufügen"
            variant="secondary"
            onPress={() => void onAddPhoto()}
            loading={photoBusy}
          />
          <PrimaryButton label="Speichern" onPress={() => void onSave()} loading={saving} />
        </>
      }
    >
      <View style={styles.badges}>
        <StatusBadge label={formatStatusLabel(defect.status)} tone={statusTone(defect.status)} />
        <StatusBadge label={priorityLabel(defect.priority)} tone={priorityTone(defect.priority)} />
      </View>
      <TextField label="Titel" value={title} onChangeText={setTitle} />
      <TextField
        label="Beschreibung"
        value={description}
        onChangeText={setDescription}
        multiline
        style={{ minHeight: 120, textAlignVertical: 'top' }}
      />

      <Section title={`Fotos (${photos.length})`}>
        {photos.length === 0 ? (
          <Text style={styles.empty}>Noch keine Fotos</Text>
        ) : (
          <View style={styles.gallery}>
            {photos.map((photo) => (
              <Image
                key={photo.id}
                source={{ uri: resolveDocumentUri(photo.local_path) }}
                style={styles.photo}
                accessibilityLabel={photo.filename}
              />
            ))}
          </View>
        )}
      </Section>

      <PrimaryButton
        label="Mangel löschen"
        variant="ghost"
        onPress={() => {
          Alert.alert('Löschen', 'Mangel soft-löschen?', [
            { text: 'Abbrechen', style: 'cancel' },
            {
              text: 'Löschen',
              style: 'destructive',
              onPress: () => {
                void defectRepository.softDeleteDefect(id).then(() => router.replace('/(tabs)/defects'));
              }
            }
          ]);
        }}
      />
    </Screen>
  );
}

const styles = StyleSheet.create({
  badges: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs,
    marginBottom: spacing.xs
  },
  empty: {
    ...typography.caption,
    color: colors.muted
  },
  gallery: {
    gap: spacing.sm
  },
  photo: {
    width: '100%',
    aspectRatio: 4 / 3,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.border
  }
});
