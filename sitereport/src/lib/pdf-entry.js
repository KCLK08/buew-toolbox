import { layoutPhotoCollage, normalizeEntryPhotos } from './photos.js';

export const PDF_A4_WIDTH = 595.28;
export const PDF_A4_HEIGHT = 841.89;
export const PDF_MARGIN = 36;
export const PDF_BLOCK_GAP = 16;
export const PDF_CARD_PADDING = 12;
export const PDF_TABLE_GAP = 12;
export const PDF_BADGE_HEIGHT = 18;
export const PDF_BADGE_GAP = 8;
export const PDF_LINE_HEIGHT = 14;
export const PDF_ROW_PAD = 6;
export const PDF_PHOTO_GAP = 4;
export const PDF_PHOTO_FRAME = 2;
export const PDF_MIN_CELL = 160;
export const PDF_MAX_CELL = 260;
export const PDF_HEADER_PADDING = 14;
export const PDF_HEADER_GAP = 18;
export const PDF_MIN_PHOTO_HEIGHT = 56;

export function pdfEntryBadgeText(index) {
  return `Eintrag ${Number(index) + 1}`;
}

export function pdfEntryNeedsPhotoArea(entry) {
  return normalizeEntryPhotos(entry).length > 0;
}

export function pdfPageBodyHeight() {
  return PDF_A4_HEIGHT - PDF_MARGIN * 2;
}

export function pdfContentWidth() {
  return PDF_A4_WIDTH - PDF_MARGIN * 2;
}

export function pdfPhotoAreaWidth() {
  return pdfContentWidth() - PDF_CARD_PADDING * 2;
}

export function pdfTwoUpCardBudget() {
  return (pdfPageBodyHeight() - PDF_BLOCK_GAP) / 2;
}

export function fitPdfPhotoCollage(sizes, maxImageWidth, maxImageHeight) {
  const list = Array.isArray(sizes) ? sizes.filter((item) => item && item.width > 0 && item.height > 0) : [];
  if (!list.length) {
    return { items: [], width: 0, height: 0, cols: 0, rows: 0 };
  }
  const budget = Math.max(PDF_MIN_PHOTO_HEIGHT, Number(maxImageHeight) || 0);
  const width = Math.max(1, Number(maxImageWidth) || pdfPhotoAreaWidth());
  const readable = layoutPhotoCollage(list, width, budget, {
    gap: PDF_PHOTO_GAP,
    frame: PDF_PHOTO_FRAME,
    minCell: Math.min(PDF_MIN_CELL, budget),
    maxCell: Math.min(PDF_MAX_CELL, budget)
  });
  if (readable.height <= budget + 1) return readable;
  return layoutPhotoCollage(list, width, budget, {
    gap: PDF_PHOTO_GAP,
    frame: PDF_PHOTO_FRAME
  });
}

export function naturalPdfPhotoCollage(sizes) {
  const list = Array.isArray(sizes) ? sizes.filter((item) => item && item.width > 0 && item.height > 0) : [];
  if (!list.length) {
    return { items: [], width: 0, height: 0, cols: 0, rows: 0 };
  }
  return layoutPhotoCollage(list, pdfPhotoAreaWidth(), pdfPageBodyHeight(), {
    gap: PDF_PHOTO_GAP,
    frame: PDF_PHOTO_FRAME,
    minCell: PDF_MIN_CELL,
    maxCell: PDF_MAX_CELL
  });
}

export function estimatePdfFieldTableHeight(entry, tableColumns) {
  const columns = Array.isArray(tableColumns) ? tableColumns : [];
  let tableHeight = 0;
  for (const col of columns) {
    const labelLines = estimateWrappedLines(col?.name || '', 24);
    const valueLines = estimateWrappedLines(entry?.fields?.[col?.name] ?? '', 52);
    const lines = Math.max(labelLines, valueLines, 1);
    tableHeight += lines * PDF_LINE_HEIGHT + PDF_ROW_PAD;
  }
  return tableHeight;
}

export function estimatePdfEntryMetrics(entry, tableColumns, photoSizes = []) {
  const sizes = Array.isArray(photoSizes)
    ? photoSizes.filter((item) => item && item.width > 0 && item.height > 0)
    : [];
  const hasImages = sizes.length > 0;
  const tableHeight = estimatePdfFieldTableHeight(entry, tableColumns);
  const chromeHeight =
    PDF_CARD_PADDING * 2 +
    PDF_BADGE_HEIGHT +
    PDF_BADGE_GAP +
    (hasImages ? PDF_TABLE_GAP : 0) +
    tableHeight;
  const collageHeight = hasImages ? naturalPdfPhotoCollage(sizes).height : 0;
  return {
    photoCount: sizes.length,
    chromeHeight,
    naturalCardHeight: chromeHeight + collageHeight
  };
}

export function estimatePdfHeaderRemaining({
  protocolTitle = '',
  protocolDescription = '',
  attendees = '',
  hasLogo = false
} = {}) {
  const titleLines = Math.max(1, estimateWrappedLines(protocolTitle || 'Protokoll', 60));
  const descLines = Math.max(1, estimateWrappedLines(protocolDescription || '—', 72, 10));
  const attendeeLines = Math.max(1, estimateWrappedLines(attendees || '—', 72, 10));
  const metaHeight =
    PDF_LINE_HEIGHT +
    PDF_LINE_HEIGHT +
    PDF_LINE_HEIGHT +
    descLines * PDF_LINE_HEIGHT +
    PDF_LINE_HEIGHT +
    attendeeLines * PDF_LINE_HEIGHT;
  const headerTextHeight = titleLines * 18 + 8 + metaHeight;
  const logoHeight = hasLogo ? 60 : 0;
  const headerBoxHeight = Math.max(headerTextHeight, logoHeight) + PDF_HEADER_PADDING * 2;
  return pdfPageBodyHeight() - headerBoxHeight - PDF_HEADER_GAP;
}

/**
 * Walk entries with the same page-break rules as the PDF export.
 */
export function layoutPdfEntryFlow({
  entries,
  tableColumns,
  photoSizesForEntry,
  headerRemaining
} = {}) {
  const list = Array.isArray(entries) ? entries : [];
  const columns = Array.isArray(tableColumns) ? tableColumns : [];
  const pageBodyHeight = pdfPageBodyHeight();
  let remaining = Number.isFinite(headerRemaining) ? headerRemaining : pageBodyHeight;
  let isFirstDocumentPage = true;
  let entriesOnPage = 0;
  const planned = [];

  for (let i = 0; i < list.length; i += 1) {
    const sizes =
      typeof photoSizesForEntry === 'function' ? photoSizesForEntry(list[i], i) || [] : [];
    const metrics = estimatePdfEntryMetrics(list[i], columns, sizes);
    let pageBreakBefore = false;
    let plan = planPdfEntryPlacement({
      remaining,
      isFirstDocumentPage,
      entriesOnPage,
      photoCount: metrics.photoCount,
      naturalCardHeight: metrics.naturalCardHeight,
      chromeHeight: metrics.chromeHeight
    });
    if (!plan.stayOnPage) {
      pageBreakBefore = true;
      remaining = pageBodyHeight;
      isFirstDocumentPage = false;
      entriesOnPage = 0;
      plan = planPdfEntryPlacement({
        remaining,
        isFirstDocumentPage: false,
        entriesOnPage: 0,
        photoCount: metrics.photoCount,
        naturalCardHeight: metrics.naturalCardHeight,
        chromeHeight: metrics.chromeHeight
      });
    }
    const maxCardHeight = Math.min(
      plan.maxCardHeight,
      Math.max(metrics.chromeHeight, remaining)
    );
    const maxImageHeight = metrics.photoCount
      ? Math.max(PDF_MIN_PHOTO_HEIGHT, Math.min(maxCardHeight, remaining) - metrics.chromeHeight)
      : 0;
    const cardHeight = metrics.photoCount
      ? metrics.chromeHeight + fitPdfPhotoCollage(sizes, pdfPhotoAreaWidth(), maxImageHeight).height
      : metrics.chromeHeight;
    planned.push({
      pageBreakBefore,
      maxCardHeight,
      maxImageHeight,
      chromeHeight: metrics.chromeHeight,
      photoCount: metrics.photoCount
    });
    remaining -= cardHeight + PDF_BLOCK_GAP;
    entriesOnPage += 1;
  }

  return planned;
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

function estimateWrappedLines(text, charsPerLine, maxLines = 40) {
  const width = Math.max(8, Number(charsPerLine) || 52);
  const raw = String(text ?? '');
  const paragraphs = raw.replace(/\r\n/g, '\n').split('\n');
  let lines = 0;
  for (const paragraph of paragraphs) {
    if (!paragraph) {
      lines += 1;
      continue;
    }
    lines += Math.max(1, Math.ceil(paragraph.length / width));
  }
  if (!lines) lines = 1;
  return Math.min(maxLines, lines);
}
