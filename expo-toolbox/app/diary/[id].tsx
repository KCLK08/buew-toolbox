import { useCallback, useMemo, useState } from 'react';
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
import { diaryRepository, photoRepository } from '../../src/repositories';
import { resolveDocumentUri } from '../../src/storage/fileService';
import type { DiaryEntry, Photo } from '../../src/types/offline';

export default function DiaryDetailScreen() {
  const { id } = useLocalSearchParams<{ id: string }>();
  const router = useRouter();
  const [entry, setEntry] = useState<DiaryEntry | null>(null);
  const [photos, setPhotos] = useState<Photo[]>([]);
  const [title, setTitle] = useState('');
  const [weather, setWeather] = useState('');
  const [notes, setNotes] = useState('');
  const [saving, setSaving] = useState(false);
  const [photoBusy, setPhotoBusy] = useState(false);

  const load = useCallback(async () => {
    if (!id) return;
    const item = await diaryRepository.getDiaryEntryById(id);
    setEntry(item);
    if (item) {
      setTitle(item.title);
      try {
        const payload = JSON.parse(item.payload_json || '{}') as {
          weather?: { notes?: string };
          notes?: string;
        };
        setWeather(payload.weather?.notes || '');
        setNotes(payload.notes || '');
      } catch {
        setWeather('');
        setNotes('');
      }
    }
    setPhotos(await photoRepository.getPhotos({ parent_id: id, parent_type: 'diary_entry' }));
  }, [id]);

  useFocusEffect(
    useCallback(() => {
      void load();
    }, [load])
  );

  const dateLabel = useMemo(
    () => entry?.entry_date || formatRelativeDate(entry?.created_at),
    [entry]
  );

  const onSave = async () => {
    if (!id || !title.trim()) return;
    setSaving(true);
    try {
      const payload = {
        weather: { notes: weather.trim() },
        notes: notes.trim(),
        personnel: [],
        equipment: [],
        work: []
      };
      await diaryRepository.updateDiaryEntry({
        id,
        title: title.trim(),
        payload_json: JSON.stringify(payload)
      });
      await load();
      Alert.alert('Gespeichert', 'Eintrag aktualisiert.');
    } finally {
      setSaving(false);
    }
  };

  const onAddPhoto = async () => {
    if (!id) return;
    setPhotoBusy(true);
    try {
      const photo = await captureAndSavePhoto({ parentId: id, parentType: 'diary_entry' });
      if (photo) await load();
    } finally {
      setPhotoBusy(false);
    }
  };

  if (!entry) {
    return (
      <Screen title="Bautagebuch" showBack>
        <EmptyState title="Eintrag nicht gefunden" actionLabel="Zurück" onAction={() => router.back()} />
      </Screen>
    );
  }

  return (
    <Screen
      title={dateLabel}
      subtitle={entry.title}
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
      <StatusBadge label={formatStatusLabel(entry.status)} tone={statusTone(entry.status)} />
      <TextField label="Titel" value={title} onChangeText={setTitle} />
      <TextField label="Wetter" value={weather} onChangeText={setWeather} />
      <TextField
        label="Notizen"
        value={notes}
        onChangeText={setNotes}
        multiline
        style={{ minHeight: 110, textAlignVertical: 'top' }}
      />

      <Section title={`Fotos (${photos.length})`}>
        {photos.length === 0 ? (
          <Text style={styles.empty}>Noch keine Fotos — nutze „+ Foto hinzufügen“.</Text>
        ) : (
          <View style={styles.gallery}>
            {photos.map((photo) => (
              <View key={photo.id} style={styles.photoWrap}>
                <Image
                  source={{ uri: resolveDocumentUri(photo.local_path) }}
                  style={styles.photo}
                  accessibilityLabel={photo.filename}
                />
                <Text style={styles.caption} numberOfLines={1}>
                  {photo.filename}
                </Text>
              </View>
            ))}
          </View>
        )}
      </Section>

      <PrimaryButton
        label="Eintrag löschen"
        variant="ghost"
        onPress={() => {
          Alert.alert('Löschen', 'Eintrag soft-löschen?', [
            { text: 'Abbrechen', style: 'cancel' },
            {
              text: 'Löschen',
              style: 'destructive',
              onPress: () => {
                void diaryRepository.softDeleteDiaryEntry(id).then(() => router.replace('/(tabs)/diary'));
              }
            }
          ]);
        }}
      />
    </Screen>
  );
}

const styles = StyleSheet.create({
  empty: {
    ...typography.caption,
    color: colors.muted
  },
  gallery: {
    gap: spacing.sm
  },
  photoWrap: {
    gap: spacing.xxs
  },
  photo: {
    width: '100%',
    aspectRatio: 4 / 3,
    borderRadius: spacing.cardRadius,
    backgroundColor: colors.border
  },
  caption: {
    ...typography.caption,
    color: colors.muted
  }
});
