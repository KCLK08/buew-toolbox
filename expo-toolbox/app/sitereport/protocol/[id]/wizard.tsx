import { useCallback, useEffect, useMemo, useState } from 'react';
import { Alert, StyleSheet, View } from 'react-native';
import { useLocalSearchParams, useRouter } from 'expo-router';

import { FieldStep, PhotoCaptureStep, ProtocolSummary, StatusFieldStep, WizardStep } from '../../../../src/components/sitereport';
import { PrimaryButton, Screen } from '../../../../src/components/mobile';
import { useToast } from '../../../../src/contexts/ToastContext';
import { spacing } from '../../../../src/constants/theme';
import type { SiteReportColumn, SiteReportEntry, SiteReportProtocol } from '../../../../src/native/sitereport/db/database';
import { captureProtocolPhoto, deleteEntryPhoto } from '../../../../src/native/sitereport/services/photoService';
import {
  createEntryId,
  emptyFieldsFromColumns,
  getProtocolOrThrow,
  upsertProtocolEntry
} from '../../../../src/native/sitereport/services/protocolService';

type WizardPhase = 'steps' | 'summary';

export default function EntryWizardScreen() {
  const router = useRouter();
  const { id, entryId } = useLocalSearchParams<{ id: string; entryId?: string }>();
  const { showToast } = useToast();
  const editingEntryId = entryId && entryId !== 'new' ? entryId : null;
  const [draftEntryId] = useState(() => editingEntryId ?? createEntryId());

  const [protocol, setProtocol] = useState<SiteReportProtocol | null>(null);
  const [stepIndex, setStepIndex] = useState(0);
  const [phase, setPhase] = useState<WizardPhase>('steps');
  const [photoPath, setPhotoPath] = useState<string | null>(null);
  const [fields, setFields] = useState<Record<string, string | number>>({});
  const [busy, setBusy] = useState(false);
  const [saving, setSaving] = useState(false);

  const steps = useMemo(() => protocol?.columns ?? [], [protocol]);

  const load = useCallback(async () => {
    if (!id) return;
    const next = await getProtocolOrThrow(id);
    setProtocol(next);
    if (editingEntryId) {
      const entry = next.entries.find((row) => row.id === editingEntryId);
      if (entry) {
        setPhotoPath(entry.photoPath);
        setFields({ ...entry.fields });
      }
    } else {
      setFields(emptyFieldsFromColumns(next.columns));
    }
  }, [editingEntryId, id]);

  useEffect(() => {
    void load();
  }, [load]);

  const currentStep: SiteReportColumn | 'summary' | null = useMemo(() => {
    if (phase === 'summary') return 'summary';
    return steps[stepIndex] ?? null;
  }, [phase, stepIndex, steps]);

  const totalSteps = steps.length + 1;

  const displayStep = phase === 'summary' ? steps.length + 1 : stepIndex + 1;

  const capturePhoto = async () => {
    if (!protocol) return;
    setBusy(true);
    try {
      const entryKey = draftEntryId;
      if (photoPath) {
        await deleteEntryPhoto(photoPath);
      }
      const path = await captureProtocolPhoto(protocol.id, entryKey);
      if (path) setPhotoPath(path);
    } catch (err) {
      Alert.alert('Kamera', err instanceof Error ? err.message : 'Foto fehlgeschlagen.');
    } finally {
      setBusy(false);
    }
  };

  const removePhoto = async () => {
    await deleteEntryPhoto(photoPath);
    setPhotoPath(null);
  };

  const goNext = () => {
    const step = steps[stepIndex];
    if (step?.isPhoto && !photoPath) {
      Alert.alert('Foto', 'Bitte zuerst ein Foto aufnehmen.');
      return;
    }
    if (stepIndex < steps.length - 1) {
      setStepIndex((prev) => prev + 1);
      return;
    }
    setPhase('summary');
  };

  const goBack = () => {
    if (phase === 'summary') {
      setPhase('steps');
      setStepIndex(steps.length - 1);
      return;
    }
    if (stepIndex > 0) {
      setStepIndex((prev) => prev - 1);
    }
  };

  const save = async () => {
    if (!protocol) return;
    setSaving(true);
    try {
      const entry: SiteReportEntry = {
        id: draftEntryId,
        createdAt: editingEntryId
          ? protocol.entries.find((row) => row.id === editingEntryId)?.createdAt ?? new Date().toISOString()
          : new Date().toISOString(),
        fields,
        photoPath
      };
      await upsertProtocolEntry(protocol, entry, editingEntryId);
      showToast(editingEntryId ? 'Eintrag aktualisiert' : 'Eintrag gespeichert');
      router.back();
    } catch (err) {
      Alert.alert('Speichern', err instanceof Error ? err.message : 'Speichern fehlgeschlagen.');
    } finally {
      setSaving(false);
    }
  };

  if (!protocol || !currentStep) {
    return (
      <Screen title="Eintrag" showBack>
        <PrimaryButton label="Laden…" disabled onPress={() => {}} />
      </Screen>
    );
  }

  const draftEntry: SiteReportEntry = {
    id: editingEntryId ?? 'draft',
    createdAt: new Date().toISOString(),
    fields,
    photoPath
  };

  const stepTitle =
    phase === 'summary'
      ? 'Zusammenfassung'
      : currentStep === 'summary'
        ? 'Zusammenfassung'
        : (currentStep as SiteReportColumn).isPhoto
          ? 'Foto aufnehmen'
          : (currentStep as SiteReportColumn).name;

  return (
    <Screen
      title={editingEntryId ? 'Eintrag bearbeiten' : 'Neuer Eintrag'}
      showBack
      footer={
        <View style={styles.footer}>
          {phase === 'steps' && stepIndex > 0 ? (
            <PrimaryButton label="Zurück" variant="secondary" onPress={goBack} />
          ) : phase === 'summary' ? (
            <PrimaryButton label="Zurück" variant="secondary" onPress={goBack} />
          ) : null}
          {phase === 'summary' ? (
            <PrimaryButton label={saving ? 'Speichern…' : 'Speichern'} disabled={saving} onPress={() => void save()} />
          ) : (
            <PrimaryButton label="Weiter" onPress={goNext} />
          )}
        </View>
      }
    >
      <WizardStep step={displayStep} total={totalSteps} title={stepTitle}>
        {phase === 'summary' ? (
          <ProtocolSummary entry={draftEntry} columns={protocol.columns} />
        ) : (currentStep as SiteReportColumn).isPhoto ? (
          <PhotoCaptureStep
            photoUri={photoPath}
            onCapture={() => void capturePhoto()}
            onRemove={() => void removePhoto()}
            busy={busy}
          />
        ) : (currentStep as SiteReportColumn).name.toLowerCase() === 'status' ? (
          <StatusFieldStep
            value={String(fields[(currentStep as SiteReportColumn).name] ?? 'offen')}
            onChange={(value) =>
              setFields((prev) => ({ ...prev, [(currentStep as SiteReportColumn).name]: value }))
            }
          />
        ) : (
          <FieldStep
            column={currentStep as SiteReportColumn}
            value={String(fields[(currentStep as SiteReportColumn).name] ?? '')}
            onChange={(value) =>
              setFields((prev) => ({
                ...prev,
                [(currentStep as SiteReportColumn).name]:
                  (currentStep as SiteReportColumn).type === 'number' && value ? Number(value) : value
              }))
            }
          />
        )}
      </WizardStep>
    </Screen>
  );
}

const styles = StyleSheet.create({
  footer: { gap: spacing.xs }
});
