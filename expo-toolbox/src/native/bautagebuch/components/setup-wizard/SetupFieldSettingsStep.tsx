import { useEffect, useMemo, useState } from 'react';
import { Pressable, ScrollView, StyleSheet, Text, View, useWindowDimensions } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { SingleLineText } from '../../../../components/mobile';
import { colors, spacing, typography } from '../../../../constants/theme';
import { systemBottomInset } from '../../../../navigation/systemInsets';
import { listSetupSections, updateSetupField } from '../../lib/setup-mapping';
import {
  listOrderedSections,
  moveSectionInSetupModel,
  sectionEntryKey
} from '../../lib/setup-section-order';
import type { DetectedField, SetupFieldConfig } from '../../types';
import { SetupPdfFieldPreview } from '../SetupPdfFieldPreview';
import { SetupFieldCard } from './SetupFieldCard';
import { SetupFieldsIntro } from './SetupFieldsIntro';
import { SetupGroupNav } from './SetupGroupNav';
import { SetupGroupPickerSheet } from './SetupGroupPickerSheet';
import { SetupSectionOrderCard } from './SetupSectionOrderCard';
import { SetupValidationList } from './SetupValidationList';

type SetupPreviewField = {
  fieldId: string;
  label: string;
  page: number;
};

type Props = {
  templateName: string;
  pdfPath: string | null;
  detectedFields: DetectedField[];
  setupModel: Record<string, unknown>;
  validationIssues: string[];
  readOnly?: boolean;
  showPreview?: boolean;
  onActiveFieldChange?: (field: SetupPreviewField | null) => void;
  onChange: (next: Record<string, unknown>) => void;
};

export function SetupFieldSettingsStep({
  templateName,
  pdfPath,
  detectedFields,
  setupModel,
  validationIssues,
  readOnly = false,
  showPreview = false,
  onActiveFieldChange,
  onChange
}: Props) {
  const insets = useSafeAreaInsets();
  const { width } = useWindowDimensions();
  const isTablet = width >= 900;
  const sections = useMemo(() => listSetupSections(setupModel), [setupModel]);
  const orderedSections = useMemo(() => listOrderedSections(setupModel), [setupModel]);
  const [activeSectionId, setActiveSectionId] = useState(sections[0]?.sectionId || '');
  const [activeFieldId, setActiveFieldId] = useState<string | null>(
    sections[0]?.fields[0]?.fieldId || null
  );
  const [groupSheetOpen, setGroupSheetOpen] = useState(false);
  const [fieldPreviewPinned, setFieldPreviewPinned] = useState(false);

  const activeSection = sections.find((section) => section.sectionId === activeSectionId) || sections[0];
  const activeField =
    activeSection?.fields.find((field) => field.fieldId === activeFieldId) ||
    activeSection?.fields[0] ||
    null;

  const fieldCount = sections.reduce((sum, section) => sum + section.fields.length, 0);
  const groupItems = sections.map((section) => ({
    sectionId: section.sectionId,
    label: section.label,
    count: section.fields.length
  }));

  useEffect(() => {
    if (sections.length === 0) return;
    const sectionExists = sections.some((section) => section.sectionId === activeSectionId);
    if (!sectionExists) {
      setActiveSectionId(sections[0].sectionId);
      setActiveFieldId(sections[0].fields[0]?.fieldId || null);
    }
  }, [sections, activeSectionId]);

  useEffect(() => {
    if (!onActiveFieldChange) return;
    if (!activeFieldId) {
      onActiveFieldChange(null);
      return;
    }
    const field = activeSection?.fields.find((entry) => entry.fieldId === activeFieldId);
    if (!field) {
      onActiveFieldChange(null);
      return;
    }
    onActiveFieldChange({
      fieldId: field.fieldId,
      label: field.label || field.fieldName || field.fieldId,
      page: field.page || 1
    });
  }, [activeFieldId, activeSection, onActiveFieldChange]);

  const handleFieldPress = (fieldId: string, expanded: boolean) => {
    const nextExpanded = !expanded;
    setActiveFieldId(nextExpanded ? fieldId : null);
    if (nextExpanded && pdfPath) {
      setFieldPreviewPinned(true);
    }
  };

  const handleFieldChange = (field: SetupFieldConfig, patch: Partial<SetupFieldConfig>) => {
    setActiveFieldId(field.fieldId);
    setFieldPreviewPinned(true);
    onChange(updateSetupField(setupModel, activeSection?.sectionId || '', field.fieldId, patch));
  };

  const selectSection = (sectionId: string) => {
    setActiveSectionId(sectionId);
    const section = sections.find((entry) => entry.sectionId === sectionId);
    setActiveFieldId(section?.fields[0]?.fieldId || null);
  };

  const activeGroupLabel =
    groupItems.find((group) => group.sectionId === activeSection?.sectionId)?.label || 'Gruppe';

  return (
    <View style={styles.root}>
      <View style={[styles.body, isTablet ? styles.bodyTablet : null]}>
        {isTablet ? (
          <View style={styles.sideNav}>
            <SetupGroupNav
              groups={groupItems}
              activeSectionId={activeSection?.sectionId || ''}
              horizontal={false}
              onSelect={selectSection}
            />
          </View>
        ) : (
          <View style={styles.mobileGroupBar}>
            <Pressable style={styles.groupPickerBtn} onPress={() => setGroupSheetOpen(true)}>
              <Text style={styles.groupPickerLabel}>Gruppe auswählen</Text>
              <SingleLineText style={styles.groupPickerValue}>{activeGroupLabel}</SingleLineText>
            </Pressable>
          </View>
        )}

        <ScrollView
          style={styles.fieldScroll}
          contentContainerStyle={[
            styles.fieldScrollContent,
            { paddingBottom: systemBottomInset(insets) + spacing.xl }
          ]}
          keyboardShouldPersistTaps="handled"
        >
          <SetupFieldsIntro
            templateName={templateName}
            fieldCount={fieldCount}
            groupCount={sections.length}
          />

          <SetupSectionOrderCard
            sections={orderedSections}
            selectedKey={
              activeSection ? sectionEntryKey({ kind: 'single', id: activeSection.sectionId }) : null
            }
            readOnly={readOnly}
            onMove={(index, direction) =>
              onChange(moveSectionInSetupModel(setupModel, index, direction))
            }
            onSelect={(entry) => {
              if (entry.kind === 'single') selectSection(entry.id);
            }}
          />

          {fieldPreviewPinned && pdfPath && activeField && !showPreview ? (
            <View style={styles.pinnedPreview}>
              <SetupPdfFieldPreview
                variant="pinned"
                emphasizeActiveHighlight
                pdfPath={pdfPath}
                detectedFields={detectedFields}
                activeFieldId={activeField.fieldId}
                activeFieldLabel={activeField.label || activeField.fieldName || null}
                activeFieldPage={activeField.page || 1}
              />
            </View>
          ) : null}

          {activeSection ? (
            <Text style={styles.sectionTitle}>{activeSection.label}</Text>
          ) : null}

          {(activeSection?.fields || []).map((field) => {
            const expanded = activeFieldId === field.fieldId;
            return (
              <SetupFieldCard
                key={field.fieldId}
                field={field}
                detectedFields={detectedFields}
                expanded={expanded}
                readOnly={readOnly}
                onPress={() => handleFieldPress(field.fieldId, expanded)}
                onChange={(patch) => handleFieldChange(field, patch)}
              />
            );
          })}

          <SetupValidationList
            issues={validationIssues}
            onSelectIssue={() => {
              // Validation messages are textual; user can scroll and fix manually.
            }}
          />
        </ScrollView>
      </View>

      <SetupGroupPickerSheet
        visible={groupSheetOpen}
        groups={groupItems}
        activeSectionId={activeSection?.sectionId || ''}
        onSelect={selectSection}
        onClose={() => setGroupSheetOpen(false)}
      />

    </View>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    position: 'relative'
  },
  body: {
    flex: 1
  },
  bodyTablet: {
    flexDirection: 'row'
  },
  sideNav: {
    width: 220,
    borderRightWidth: 1,
    borderRightColor: colors.border
  },
  mobileGroupBar: {
    paddingHorizontal: spacing.pageX,
    paddingVertical: spacing.sm,
    borderBottomWidth: 1,
    borderBottomColor: colors.border,
    backgroundColor: colors.panel
  },
  groupPickerBtn: {
    minHeight: spacing.touchMin + 4,
    borderRadius: 12,
    borderWidth: 1,
    borderColor: colors.border,
    backgroundColor: colors.panelElevated,
    paddingHorizontal: spacing.md,
    paddingVertical: spacing.sm,
    gap: 2,
    justifyContent: 'center'
  },
  groupPickerLabel: {
    ...typography.caption,
    color: colors.muted
  },
  groupPickerValue: {
    ...typography.bodyStrong,
    color: colors.ink
  },
  fieldScroll: {
    flex: 1
  },
  fieldScrollContent: {
    gap: spacing.sm,
    padding: spacing.pageX
  },
  sectionTitle: {
    ...typography.subtitle,
    color: colors.ink
  },
  pinnedPreview: {
    minHeight: 220,
    borderRadius: 12,
    overflow: 'hidden',
    borderWidth: 1,
    borderColor: colors.border
  }
});
