import { useEffect, useState } from 'react';
import { Alert, Pressable, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { PrimaryButton, Screen, TextField } from '../../src/components/mobile';
import { colors, spacing, typography } from '../../src/constants/theme';
import { diaryRepository, projectRepository } from '../../src/repositories';
import type { Project } from '../../src/types/offline';

export default function NewDiaryScreen() {
  const router = useRouter();
  const [title, setTitle] = useState('');
  const [notes, setNotes] = useState('');
  const [weather, setWeather] = useState('');
  const [projects, setProjects] = useState<Project[]>([]);
  const [projectId, setProjectId] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    void projectRepository.getProjects().then((rows) => {
      setProjects(rows);
      if (rows[0]) setProjectId(rows[0].id);
    });
  }, []);

  const onSave = async () => {
    if (!title.trim()) {
      Alert.alert('Bautagebuch', 'Bitte einen Titel angeben.');
      return;
    }
    setSaving(true);
    try {
      const payload = {
        weather: { notes: weather.trim() },
        notes: notes.trim(),
        personnel: [],
        equipment: [],
        work: []
      };
      const entry = await diaryRepository.createDiaryEntry({
        title: title.trim(),
        project_id: projectId,
        entry_date: new Date().toISOString().slice(0, 10),
        status: 'draft',
        payload_json: JSON.stringify(payload)
      });
      router.replace(`/diary/${entry.id}`);
    } catch (error) {
      Alert.alert('Fehler', error instanceof Error ? error.message : 'Speichern fehlgeschlagen.');
    } finally {
      setSaving(false);
    }
  };

  return (
    <Screen
      title="Neuer Eintrag"
      showBack
      footer={<PrimaryButton label="Eintrag speichern" onPress={() => void onSave()} loading={saving} />}
    >
      <TextField label="Titel" value={title} onChangeText={setTitle} placeholder="z. B. Gleisarbeiten" autoFocus />

      <Text style={styles.label}>Projekt</Text>
      <View style={styles.chips}>
        <Pressable
          onPress={() => setProjectId(null)}
          style={[styles.chip, projectId === null ? styles.chipActive : null]}
        >
          <Text style={[styles.chipLabel, projectId === null ? styles.chipLabelActive : null]}>Kein Projekt</Text>
        </Pressable>
        {projects.map((project) => (
          <Pressable
            key={project.id}
            onPress={() => setProjectId(project.id)}
            style={[styles.chip, projectId === project.id ? styles.chipActive : null]}
          >
            <Text style={[styles.chipLabel, projectId === project.id ? styles.chipLabelActive : null]}>
              {project.name}
            </Text>
          </Pressable>
        ))}
      </View>

      <TextField label="Wetter" value={weather} onChangeText={setWeather} placeholder="z. B. bedeckt, 12°C" />
      <TextField
        label="Notizen"
        value={notes}
        onChangeText={setNotes}
        placeholder="Kurznotiz zum Bautag"
        multiline
        style={{ minHeight: 110, textAlignVertical: 'top' }}
      />
    </Screen>
  );
}

const styles = StyleSheet.create({
  label: {
    ...typography.label,
    color: colors.ink,
    marginBottom: spacing.xs
  },
  chips: {
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: spacing.xs,
    marginBottom: spacing.md
  },
  chip: {
    minHeight: spacing.touchMin,
    paddingHorizontal: spacing.md,
    borderRadius: 999,
    borderWidth: 1,
    borderColor: colors.borderStrong,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.panel
  },
  chipActive: {
    backgroundColor: colors.accent,
    borderColor: colors.accent
  },
  chipLabel: {
    ...typography.label,
    color: colors.ink
  },
  chipLabelActive: {
    color: colors.white
  }
});
