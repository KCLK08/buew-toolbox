import { useEffect, useMemo, useState } from 'react';
import { StyleSheet, View } from 'react-native';

import { BottomSheet, BottomSheetOption, PrimaryButton } from '../../../components/mobile';
import { spacing } from '../../../constants/theme';
import type { BautagebuchExportMode } from '../services/exportService';

type Props = {
  visible: boolean;
  photoCount: number;
  photoDocEnabled: boolean | null | undefined;
  exporting?: boolean;
  onClose: () => void;
  onExport: (mode: BautagebuchExportMode) => void;
};

const EXPORT_OPTIONS: Array<{
  mode: BautagebuchExportMode;
  label: string;
  description: string;
}> = [
  {
    mode: 'btb',
    label: 'Nur Bautagebuch (BTB)',
    description: 'Ausgefülltes eBTB-PDF ohne separate Fotodokumentation.'
  },
  {
    mode: 'photo',
    label: 'Nur Fotodokumentation',
    description: 'PDF mit allen Fotodoku-Bildern als eigenständiges Dokument.'
  },
  {
    mode: 'merged',
    label: 'BTB + Fotodoku',
    description: 'Zusammengeführtes PDF mit Bautagebuch und Fotos in einer Datei.'
  }
];

function defaultMode(photoCount: number, photoDocEnabled: boolean | null | undefined): BautagebuchExportMode {
  if (photoDocEnabled === true || photoCount > 0) {
    return 'merged';
  }
  return 'btb';
}

function exportButtonLabel(mode: BautagebuchExportMode, exporting: boolean): string {
  if (exporting) return 'PDF wird erstellt…';
  if (mode === 'btb') return 'BTB exportieren & teilen';
  if (mode === 'photo') return 'Fotodoku exportieren & teilen';
  return 'BTB + Fotodoku exportieren & teilen';
}

export function ExportFinishSheet({
  visible,
  photoCount,
  photoDocEnabled,
  exporting = false,
  onClose,
  onExport
}: Props) {
  const [mode, setMode] = useState<BautagebuchExportMode>('merged');

  useEffect(() => {
    if (visible) {
      setMode(defaultMode(photoCount, photoDocEnabled));
    }
  }, [visible, photoCount, photoDocEnabled]);

  const options = useMemo(() => {
    const hasPhotos = photoCount > 0;
    const photoRelevant = photoDocEnabled === true || hasPhotos;
    return EXPORT_OPTIONS.map((option) => ({
      ...option,
      disabled: !photoRelevant && option.mode !== 'btb'
    }));
  }, [photoCount, photoDocEnabled]);

  return (
    <BottomSheet
      visible={visible}
      title="BTB abschließen"
      subtitle="Wähle, welche PDF-Version erstellt und geteilt werden soll."
      onClose={onClose}
    >
      <View style={styles.options}>
        {options.map((option) => (
          <BottomSheetOption
            key={option.mode}
            label={option.label}
            description={option.description}
            selected={mode === option.mode}
            disabled={option.disabled}
            onPress={() => setMode(option.mode)}
          />
        ))}
      </View>
      <View style={styles.actions}>
        <PrimaryButton
          label={exportButtonLabel(mode, exporting)}
          disabled={exporting}
          loading={exporting}
          onPress={() => onExport(mode)}
        />
        <PrimaryButton label="Abbrechen" variant="ghost" disabled={exporting} onPress={onClose} />
      </View>
    </BottomSheet>
  );
}

const styles = StyleSheet.create({
  options: { gap: spacing.xs },
  actions: { gap: spacing.xs, marginTop: spacing.sm }
});
