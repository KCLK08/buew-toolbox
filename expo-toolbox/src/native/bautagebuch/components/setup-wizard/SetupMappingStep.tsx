import { useMemo } from 'react';
import { Pressable, StyleSheet, Text, View } from 'react-native';

import { PrimaryButton } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import {
  assignFieldToGroup,
  addWizardGroup,
  deferField,
  getMappingProgress,
  getWizardState,
  resolveOverlayPlacement,
  type MappingField
} from '../../lib/setup-mapping';
import type { DetectedField } from '../../types';
import { SetupPdfFieldPreview } from '../SetupPdfFieldPreview';
import { GroupOverlayCards } from './GroupOverlayCards';
import { SetupProgressHeader } from './SetupProgressHeader';

type Props = {
  pdfPath: string | null;
  detectedFields: DetectedField[];
  mappingFields: MappingField[];
  setupModel: Record<string, unknown>;
  onChange: (next: Record<string, unknown>) => void;
  onComplete: () => void;
  onFinishLater: () => void;
};

export function SetupMappingStep({
  pdfPath,
  detectedFields,
  mappingFields,
  setupModel,
  onChange,
  onComplete,
  onFinishLater
}: Props) {
  const wizard = useMemo(() => getWizardState(setupModel), [setupModel]);
  const currentIndex = Math.min(
    Math.max(0, wizard.currentFieldIndex),
    Math.max(0, mappingFields.length - 1)
  );
  const currentField = mappingFields[currentIndex] || null;
  const progress = useMemo(
    () => getMappingProgress(mappingFields, wizard),
    [mappingFields, wizard]
  );
  const placement = resolveOverlayPlacement(currentField?.rect || null);

  const assignGroup = (sectionId: string) => {
    if (!currentField) return;
    let next = assignFieldToGroup(setupModel, currentField.fieldId, sectionId);
    const nextWizard = getWizardState(next);
    const nextIndex = Math.min(currentIndex + 1, Math.max(0, mappingFields.length - 1));
    next = {
      ...next,
      wizard: {
        ...nextWizard,
        currentFieldIndex: nextIndex
      }
    };
    onChange(next);

    const assignedCount = mappingFields.filter((field) =>
      Boolean(getWizardState(next).assignments[field.fieldId])
    ).length;
    if (assignedCount >= mappingFields.length) {
      onComplete();
    }
  };

  const createGroup = (label: string) => {
    const result = addWizardGroup(setupModel, label);
    onChange(result.setupModel);
  };

  const goBack = () => {
    onChange({
      ...setupModel,
      wizard: {
        ...wizard,
        currentFieldIndex: Math.max(0, currentIndex - 1)
      }
    });
  };

  const goForward = () => {
    onChange({
      ...setupModel,
      wizard: {
        ...wizard,
        currentFieldIndex: Math.min(mappingFields.length - 1, currentIndex + 1)
      }
    });
  };

  const skipField = () => {
    if (!currentField) return;
    let next = deferField(setupModel, currentField.fieldId);
    next = {
      ...next,
      wizard: {
        ...getWizardState(next),
        currentFieldIndex: Math.min(mappingFields.length - 1, currentIndex + 1)
      }
    };
    onChange(next);
  };

  return (
    <View style={styles.root}>
      <SetupProgressHeader
        progress={progress}
        title={currentField ? currentField.labelCandidate : 'Feldzuordnung'}
      />

      <View style={styles.canvas}>
        <SetupPdfFieldPreview
          pdfPath={pdfPath}
          detectedFields={detectedFields}
          activeFieldId={currentField?.fieldId || null}
          activeFieldLabel={currentField?.labelCandidate || null}
          activeFieldPage={currentField?.page || 1}
          variant="mapping"
        />

        {currentField ? (
          <GroupOverlayCards
            groups={wizard.groups}
            placement={placement}
            onSelectGroup={assignGroup}
            onCreateGroup={createGroup}
          />
        ) : null}
      </View>

      <View style={styles.footer}>
        <View style={styles.navRow}>
          <Pressable style={styles.navBtn} onPress={goBack} disabled={currentIndex <= 0}>
            <Text style={styles.navLabel}>Zurück</Text>
          </Pressable>
          <Pressable style={styles.navBtn} onPress={skipField}>
            <Text style={styles.navLabel}>Überspringen</Text>
          </Pressable>
          <Pressable style={styles.navBtn} onPress={goForward}>
            <Text style={styles.navLabel}>Weiter</Text>
          </Pressable>
        </View>
        <PrimaryButton label="Später bearbeiten" variant="ghost" onPress={onFinishLater} />
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1
  },
  canvas: {
    flex: 1,
    position: 'relative'
  },
  footer: {
    gap: spacing.sm,
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.md,
    borderTopWidth: 1,
    borderTopColor: colors.border,
    backgroundColor: colors.panel
  },
  navRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    gap: spacing.sm
  },
  navBtn: {
    minHeight: spacing.touchMin,
    justifyContent: 'center',
    paddingHorizontal: spacing.sm
  },
  navLabel: {
    ...typography.bodyStrong,
    color: colors.accent2
  }
});
