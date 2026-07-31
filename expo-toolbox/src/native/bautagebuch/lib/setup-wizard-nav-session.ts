import type { SetupWizardStep } from '../types';

/** In-flight step navigation target (survives screen unmount). */
let navigatingToStep: SetupWizardStep | null = null;

export function markSetupStepNavigation(step: SetupWizardStep): void {
  navigatingToStep = step;
}

export function isSetupStepNavigationActive(): boolean {
  return navigatingToStep !== null;
}

export function getSetupStepNavigationTarget(): SetupWizardStep | null {
  return navigatingToStep;
}

/** Returns true when this screen should apply the pending navigation step. */
export function consumeSetupStepNavigation(expectedStep: SetupWizardStep): boolean {
  if (navigatingToStep !== expectedStep) {
    return false;
  }
  navigatingToStep = null;
  return true;
}

export function clearSetupStepNavigation(): void {
  navigatingToStep = null;
}
