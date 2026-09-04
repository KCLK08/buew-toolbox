/**
 * Photo helpers for SiteReport entries.
 * Supports legacy `photoBlob` plus `photoBlobs[]`.
 */

export function normalizeEntryPhotos(entry) {
  if (!entry) return [];
  if (Array.isArray(entry.photoBlobs) && entry.photoBlobs.length) {
    return entry.photoBlobs.filter(Boolean);
  }
  if (entry.photoBlob) return [entry.photoBlob];
  if (Array.isArray(entry.photoFiles) && entry.photoFiles.length) {
    return entry.photoFiles.filter(Boolean);
  }
  if (entry.photoFile) return [entry.photoFile];
  return [];
}

export function preferredPhotoColumns(count) {
  const n = Number(count) || 0;
  if (n <= 0) return 0;
  if (n <= 3) return n;
  if (n === 4) return 2;
  if (n <= 6) return 3;
  if (n <= 8) return 4;
  return Math.ceil(Math.sqrt(n));
}

/**
 * How many columns still keep each cell at least `minCell` wide.
 */
export function preferredReadableColumns(count, maxWidth, minCell, gap = 4) {
  const n = Number(count) || 0;
  if (n <= 0) return 0;
  const widthLimit = Math.max(1, Number(maxWidth) || 1);
  const cellMin = Math.max(1, Number(minCell) || 1);
  const maxCols = Math.max(1, Math.floor((widthLimit + gap) / (cellMin + gap)));
  return Math.min(n, maxCols);
}

/**
 * Pack images into a framed grid.
 *
 * Default: fit inside maxWidth x maxHeight (may shrink photos).
 * With `minCell`: keep photos readable and let the collage grow taller instead.
 *
 * @param {Array<{width:number, height:number}>} sizes
 * @param {number} maxWidth
 * @param {number} maxHeight
 * @param {{gap?: number, frame?: number, minCell?: number, maxCell?: number}} [opts]
 */
export function layoutPhotoCollage(sizes, maxWidth, maxHeight, opts = {}) {
  const itemsIn = Array.isArray(sizes) ? sizes.filter((s) => s && s.width > 0 && s.height > 0) : [];
  const n = itemsIn.length;
  const gap = Number.isFinite(opts.gap) ? opts.gap : 4;
  const frame = Number.isFinite(opts.frame) ? opts.frame : 2;
  const widthLimit = Math.max(1, Number(maxWidth) || 1);
  const heightLimit = Math.max(1, Number(maxHeight) || 1);
  const minCell = Number.isFinite(opts.minCell) && opts.minCell > 0 ? opts.minCell : 0;

  if (n === 0) {
    return { items: [], width: 0, height: 0, cols: 0, rows: 0 };
  }

  if (minCell > 0) {
    return layoutPhotoCollageReadable(itemsIn, widthLimit, {
      gap,
      frame,
      minCell,
      maxCell: opts.maxCell
    });
  }

  const candidates = new Set(
    n <= 3
      ? [n]
      : [preferredPhotoColumns(n), n, Math.ceil(n / 2), Math.min(n, 4)]
  );
  let best = null;

  for (const cols of candidates) {
    if (!cols || cols < 1) continue;
    const rows = Math.ceil(n / cols);
    const cellW = (widthLimit - gap * (cols - 1)) / cols;
    const cellH = (heightLimit - gap * (rows - 1)) / rows;
    if (cellW < 8 || cellH < 8) continue;

    const items = itemsIn.map((size, index) => {
      const col = index % cols;
      const row = Math.floor(index / cols);
      const innerW = Math.max(1, cellW - frame * 2);
      const innerH = Math.max(1, cellH - frame * 2);
      const scale = Math.min(innerW / size.width, innerH / size.height);
      const width = Math.max(1, size.width * scale);
      const height = Math.max(1, size.height * scale);
      const frameX = col * (cellW + gap);
      const frameY = row * (cellH + gap);
      return {
        x: frameX + frame + (innerW - width) / 2,
        y: frameY + frame + (innerH - height) / 2,
        width,
        height,
        frameX,
        frameY,
        frameW: cellW,
        frameH: cellH
      };
    });

    const width = cols * cellW + gap * (cols - 1);
    const height = rows * cellH + gap * (rows - 1);
    const minArea = Math.min(...items.map((item) => item.width * item.height));
    const score = minArea * (rows === 1 ? 1.1 : 1);

    if (!best || score > best.score) {
      best = { items, width, height, cols, rows, score };
    }
  }

  if (!best) {
    return { items: [], width: 0, height: 0, cols: 0, rows: 0 };
  }

  const { score, ...layout } = best;
  return layout;
}

function layoutPhotoCollageReadable(itemsIn, widthLimit, opts) {
  const n = itemsIn.length;
  const { gap, frame, minCell } = opts;
  const maxCell = Number.isFinite(opts.maxCell) && opts.maxCell > 0 ? opts.maxCell : Infinity;
  const cols = preferredReadableColumns(n, widthLimit, minCell, gap);
  const rows = Math.ceil(n / cols);
  const cellW = (widthLimit - gap * (cols - 1)) / cols;
  const innerW = Math.max(1, cellW - frame * 2);

  const rowHeights = Array.from({ length: rows }, () => minCell);
  itemsIn.forEach((size, index) => {
    const row = Math.floor(index / cols);
    const scaledH = innerW * (size.height / size.width) + frame * 2;
    rowHeights[row] = Math.min(maxCell, Math.max(rowHeights[row], scaledH));
  });

  const items = itemsIn.map((size, index) => {
    const col = index % cols;
    const row = Math.floor(index / cols);
    const cellH = rowHeights[row];
    const innerH = Math.max(1, cellH - frame * 2);
    const scale = Math.min(innerW / size.width, innerH / size.height);
    const width = Math.max(1, size.width * scale);
    const height = Math.max(1, size.height * scale);
    const frameX = col * (cellW + gap);
    const frameY = rowHeights.slice(0, row).reduce((sum, h) => sum + h + gap, 0);
    return {
      x: frameX + frame + (innerW - width) / 2,
      y: frameY + frame + (innerH - height) / 2,
      width,
      height,
      frameX,
      frameY,
      frameW: cellW,
      frameH: cellH
    };
  });

  const width = cols * cellW + gap * (cols - 1);
  const height = rowHeights.reduce((sum, h, i) => sum + h + (i > 0 ? gap : 0), 0);
  return { items, width, height, cols, rows };
}
