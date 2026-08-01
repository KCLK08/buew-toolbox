export type KeyboardMetrics = {
  screenY: number;
  height: number;
};

export function resolveKeyboardVisibleBounds(
  windowHeight: number,
  keyboard: KeyboardMetrics | null,
  footerInset: number,
  topInset = 0
): { top: number; bottom: number } {
  const bottom = keyboard ? keyboard.screenY - footerInset : windowHeight - footerInset;
  return { top: topInset, bottom };
}

export function resolveKeyboardScrollDelta(
  fieldY: number,
  fieldHeight: number,
  visibleTop: number,
  visibleBottom: number,
  padding = 28
): number {
  const fieldBottom = fieldY + fieldHeight;
  if (fieldBottom > visibleBottom - padding) {
    return fieldBottom - visibleBottom + padding;
  }
  if (fieldY < visibleTop + padding) {
    return fieldY - visibleTop - padding;
  }
  return 0;
}
