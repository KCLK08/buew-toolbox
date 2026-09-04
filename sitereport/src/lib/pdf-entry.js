import { normalizeEntryPhotos } from './photos.js';

export const PDF_A4_HEIGHT = 841.89;
export const PDF_MARGIN = 36;
export const PDF_BLOCK_GAP = 16;

export function pdfEntryBadgeText(index) {
  return `Eintrag ${Number(index) + 1}`;
}

export function pdfEntryNeedsPhotoArea(entry) {
  return normalizeEntryPhotos(entry).length > 0;
}

export function pdfPageBodyHeight() {
  return PDF_A4_HEIGHT - PDF_MARGIN * 2;
}

export function pdfTwoUpCardBudget() {
  return (pdfPageBodyHeight() - PDF_BLOCK_GAP) / 2;
}

/**
 * Decide whether the next entry stays on the current page and how tall it may be.
 * First page: always keep the first entry under the header.
 * Later pages: try to place two entries, unless many photos need a full page.
 */
export function planPdfEntryPlacement({
  remaining,
  isFirstDocumentPage,
  entriesOnPage,
  photoCount,
  naturalCardHeight,
  chromeHeight
}) {
  const pageBodyHeight = pdfPageBodyHeight();
  const twoUpMax = pdfTwoUpCardBudget();
  const space = Math.max(0, Number(remaining) || 0);
  const onPage = Number(entriesOnPage) || 0;
  const photos = Number(photoCount) || 0;
  const natural = Math.max(0, Number(naturalCardHeight) || 0);
  const chrome = Math.max(0, Number(chromeHeight) || 0);
  const manyPhotos = photos >= 4 || (photos >= 3 && natural > twoUpMax + 4);
  const minCard = chrome + (photos ? 72 : 0);

  if (isFirstDocumentPage && onPage === 0) {
    return {
      stayOnPage: true,
      maxCardHeight: Math.max(minCard, space),
      forceFit: true
    };
  }

  if (onPage >= 2) {
    return {
      stayOnPage: false,
      maxCardHeight: manyPhotos ? pageBodyHeight : twoUpMax,
      forceFit: false
    };
  }

  if (isFirstDocumentPage && onPage === 1) {
    if (natural <= space) {
      return { stayOnPage: true, maxCardHeight: space, forceFit: false };
    }
    return {
      stayOnPage: false,
      maxCardHeight: manyPhotos ? pageBodyHeight : twoUpMax,
      forceFit: false
    };
  }

  if (onPage === 0) {
    if (manyPhotos) {
      return { stayOnPage: true, maxCardHeight: space, forceFit: false };
    }
    return {
      stayOnPage: true,
      maxCardHeight: Math.min(space, twoUpMax),
      forceFit: false
    };
  }

  if (natural <= space) {
    return { stayOnPage: true, maxCardHeight: space, forceFit: false };
  }
  if (!manyPhotos && space >= minCard) {
    return { stayOnPage: true, maxCardHeight: space, forceFit: true };
  }
  return {
    stayOnPage: false,
    maxCardHeight: manyPhotos ? pageBodyHeight : twoUpMax,
    forceFit: false
  };
}
