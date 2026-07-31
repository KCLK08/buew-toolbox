import type { Href } from 'expo-router';

import type { SetupWizardStep } from '../types';
import { resolveSetupEditStepPath, withWizardState } from './setup-mapping';

/** Vorlagen-Setup Übersicht (Tab „Vorlagen-Setup“). */
export const SETUP_OVERVIEW_PATH = '/bautagebuch/config' as Href;

export function resolveSetupOverviewPath(): Href {
  return SETUP_OVERVIEW_PATH;
}

export function prepareSetupStepNavigation(
  setupModel: Record<string, unknown>,
  step: SetupWizardStep
): Record<string, unknown> {
  return withWizardState(setupModel, { step });
}

export function resolveSetupStepPath(templateId: string, step: SetupWizardStep): Href {
  return resolveSetupEditStepPath(templateId, step);
}
