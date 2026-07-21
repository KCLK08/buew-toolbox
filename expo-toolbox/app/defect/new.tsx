import { useState } from 'react';
import { Alert, Pressable, StyleSheet, Text, View } from 'react-native';
import { useRouter } from 'expo-router';

import { PrimaryButton, Screen, TextField } from '../../src/components/mobile';
import { colors, spacing, typography } from '../../src/constants/theme';
import { defectRepository } from '../../src/repositories';
import type { DefectPriority } from '../../src/types/offline';

const PRIORITIES: DefectPriority[] = ['low', 'normal', 'high', 'critical'];

export default function NewDefectScreen() {
  const router = useRouter();
  const [title, setTitle] = useState('');
  const [description, setDescription] = useState('');
  const [priority, setPriority] = useState<DefectPriority>('normal');
  const [saving, setSaving] = useState(false);

  const onSave = async () => {
    if (!title.trim()) {
      Alert.alert('Mangel', 'Bitte einen Titel angeben.');
      return;
    }
    setSaving(true);
    try {
      const defect = await defectRepository.createDefect({
        title: title.trim(),
        description: description.trim(),
        priority,
        status: 'draft'
      });
      router.replace(`/defect/${defect.id}`);
    } catch (error) {
      Alert.alert('Fehler', error instanceof Error ? error.message : 'Speichern fehlgeschlagen.');
    } finally {
      setSaving(false);
    }
  };

  return (
    <Screen
      title="Neuer Mangel"
      showBack
      footer={<PrimaryButton label="Mangel speichern" onPress={() => void onSave()} loading={saving} />}
    >
      <TextField label="Titel" value={title} onChangeText={setTitle} placeholder="Kurzbeschreibung" autoFocus />
      <Text style={styles.label}>Priorität</Text>
      <View style={styles.priorityRow}>
        {PRIORITIES.map((item) => (
          <Pressable
            key={item}
            onPress={() => setPriority(item)}
            style={[styles.chip, priority === item ? styles.chipActive : null]}
          >
            <Text style={[styles.chipLabel, priority === item ? styles.chipLabelActive : null]}>{item}</Text>
          </Pressable>
        ))}
      </View>
      <TextField
        label="Beschreibung"
        value={description}
        onChangeText={setDescription}
        placeholder="Details zum Mangel"
        multiline
        style={{ minHeight: 120, textAlignVertical: 'top' }}
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
  priorityRow: {
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
    color: colors.ink,
    textTransform: 'capitalize'
  },
  chipLabelActive: {
    color: colors.white
  }
});
