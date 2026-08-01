import { Alert } from 'react-native';
import type { Router } from 'expo-router';

import {
  prepareSetupStepNavigation,
  resolveSetupOverviewPath,
  resolveSetupStepPath
} from '../lib/setup-wizard-navigation';
import { markSetupStepNavigation } from '../lib/setup-wizard-nav-session';
import type { SetupWizardStep } from '../types';

type AutosaveApi = {
  schedule: (model: Record<string, unknown>) => void;
  flush: () => Promise<void>;
  isPending: () => boolean;
};

export async function flushSetupAutosave(autosave: AutosaveApi): Promise<void> {
  await autosave.flush();
}

export async function navigateSetupWizardStep(options: {
  templateId: string;
  step: SetupWizardStep;
  setupModel: Record<string, unknown>;
  autosave: AutosaveApi;
  setSetupModel: (next: Record<string, unknown>) => void;
  router: Pick<Router, 'replace'>;
}): Promise<void> {
  markSetupStepNavigation(options.step);
  const next = prepareSetupStepNavigation(options.setupModel, options.step);
  options.setSetupModel(next);
  options.autosave.schedule(next);
  await flushSetupAutosave(options.autosave);
  options.router.replace(resolveSetupStepPath(options.templateId, options.step));
}

export function confirmSetupWizardExit(options: {
  autosave: AutosaveApi;
  router: Pick<Router, 'replace'>;
  onComplete?: () => void;
}): void {
  const leave = () => {
    void (async () => {
      await flushSetupAutosave(options.autosave);
      options.onComplete?.();
      options.router.replace(resolveSetupOverviewPath());
    })();
  };

  if (!options.autosave.isPending()) {
    leave();
    return;
  }

  Alert.alert('Setup verlassen', 'Änderungen speichern?', [
    { text: 'Abbrechen', style: 'cancel' },
    {
      text: 'Verwerfen',
      style: 'destructive',
      onPress: () => {
        options.router.replace(resolveSetupOverviewPath());
      }
    },
    {
      text: 'Speichern und verlassen',
      onPress: leave
    }
  ]);
}

export async function exitSetupWizardToOverview(options: {
  autosave: AutosaveApi;
  router: Pick<Router, 'replace'>;
}): Promise<void> {
  await flushSetupAutosave(options.autosave);
  options.router.replace(resolveSetupOverviewPath());
}
