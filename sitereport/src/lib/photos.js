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
 * Pack images into a framed grid that fits inside maxWidth x maxHeight.
 * Prefers a single row for few images, then wraps so the collage stays on one page.
 *
 * @param {Array<{width:number, height:number}>} sizes
 * @param {number} maxWidth
 * @param {number} maxHeight
 * @param {{gap?: number, frame?: number}} [opts]
 */
export function layoutPhotoCollage(sizes, maxWidth, maxHeight, opts = {}) {
  const itemsIn = Array.isArray(sizes) ? sizes.filter((s) => s && s.width > 0 && s.height > 0) : [];
  const n = itemsIn.length;
  const gap = Number.isFinite(opts.gap) ? opts.gap : 4;
  const frame = Number.isFinite(opts.frame) ? opts.frame : 2;
  const widthLimit = Math.max(1, Number(maxWidth) || 1);
  const heightLimit = Math.max(1, Number(maxHeight) || 1);

  if (n === 0) {
    return { items: [], width: 0, height: 0, cols: 0, rows: 0 };
  }

  const candidates = new Set([preferredPhotoColumns(n), n, Math.ceil(n / 2), Math.min(n, 4)]);
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
    const score = minArea * (n <= 3 && rows === 1 ? 1.15 : 1);

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
