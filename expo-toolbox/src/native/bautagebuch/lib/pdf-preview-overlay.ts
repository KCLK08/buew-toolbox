/** PDF page dimensions (A4-like, PDF points, origin bottom-left). */
const PDF_PAGE_HEIGHT = 842;
const PDF_PAGE_WIDTH = 595;

export type PreviewOverlayPlacement = 'top' | 'bottom' | 'left' | 'right' | 'bottom-sheet';

/**
 * Preview-side overlay placement — mirrors setup-mapping.resolveOverlayPlacement
 * with bottom-sheet fallback when the field occupies the vertical center band.
 */
export function resolvePreviewOverlayPlacement(rect: number[] | null): PreviewOverlayPlacement {
  if (!rect || rect.length < 4) return 'bottom-sheet';

  const [x1, y1, x2, y2] = rect;
  const centerX = (x1 + x2) / 2;
  const centerY = (y1 + y2) / 2;

  const topThird = PDF_PAGE_HEIGHT * 0.68;
  const bottomThird = PDF_PAGE_HEIGHT * 0.32;
  const leftThird = PDF_PAGE_WIDTH * 0.33;
  const rightThird = PDF_PAGE_WIDTH * 0.67;
  const midBandTop = PDF_PAGE_HEIGHT * 0.58;
  const midBandBottom = PDF_PAGE_HEIGHT * 0.42;

  if (centerY >= topThird) return 'bottom';
  if (centerY <= bottomThird) return 'top';
  if (centerX <= leftThird) return 'right';
  if (centerX >= rightThird) return 'left';

  if (centerY > midBandBottom && centerY < midBandTop) {
    return 'bottom-sheet';
  }

  return 'bottom';
}

/** Scroll bias passed into WebView for ergonomic field positioning. */
export function previewScrollOverlayPlacement(
  placement: PreviewOverlayPlacement
): 'top' | 'bottom' | 'left' | 'right' {
  if (placement === 'bottom-sheet') return 'bottom';
  return placement;
}
