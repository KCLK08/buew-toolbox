import { useState } from 'react';
import { Alert } from 'react-native';
import { useRouter } from 'expo-router';

import { PrimaryButton, Screen, TextField } from '../../src/components/mobile';
import { projectRepository } from '../../src/repositories';

export default function NewProjectScreen() {
  const router = useRouter();
  const [name, setName] = useState('');
  const [location, setLocation] = useState('');
  const [description, setDescription] = useState('');
  const [saving, setSaving] = useState(false);

  const onSave = async () => {
    if (!name.trim()) {
      Alert.alert('Projekt', 'Bitte einen Namen angeben.');
      return;
    }
    setSaving(true);
    try {
      const project = await projectRepository.createProject({
        name: name.trim(),
        location: location.trim(),
        description: description.trim(),
        status: 'active',
        date: new Date().toISOString().slice(0, 10)
      });
      router.replace(`/project/${project.id}`);
    } catch (error) {
      Alert.alert('Fehler', error instanceof Error ? error.message : 'Speichern fehlgeschlagen.');
    } finally {
      setSaving(false);
    }
  };

  return (
    <Screen
      title="Neues Projekt"
      showBack
      footer={<PrimaryButton label="Projekt speichern" onPress={() => void onSave()} loading={saving} />}
    >
      <TextField label="Name" value={name} onChangeText={setName} placeholder="Baustelle / Projekt" autoFocus />
      <TextField label="Ort" value={location} onChangeText={setLocation} placeholder="Adresse oder Ort" />
      <TextField
        label="Beschreibung"
        value={description}
        onChangeText={setDescription}
        placeholder="Kurzbeschreibung"
        multiline
        style={{ minHeight: 96, textAlignVertical: 'top' }}
      />
    </Screen>
  );
}
