import { normalizeEntryPhotos } from './photos.js';

export function pdfEntryBadgeText(index) {
  return `Eintrag ${Number(index) + 1}`;
}

export function pdfEntryNeedsPhotoArea(entry) {
  return normalizeEntryPhotos(entry).length > 0;
}
