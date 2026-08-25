const MIN_READABLE_FONT_SIZE = 6;
const DEFAULT_START_FONT_SIZE = 12;
const LINE_HEIGHT_FACTOR = 1.22;
const CHAR_WIDTH_FACTOR = 0.52;

const AUTO_FIT_TABLE_COLUMNS = new Set(['table_detail_blocks:c4', 'table_detail_blocks:c5']);
const AUTO_FIT_FIELD_NAMES = new Set(['Text66', 'Text67']);

export function normalizeMultilineText(text) {
  return String(text ?? '').replace(/\r\n/g, '\n');
}

export function shouldAutoFitPdfField({ tableId = '', columnId = '', fieldName = '', autoFit = false } = {}) {
  if (autoFit === true) {
    return true;
  }

  const normalizedTableId = String(tableId || '').trim();
  const normalizedColumnId = String(columnId || '').trim();
  if (normalizedTableId && normalizedColumnId && AUTO_FIT_TABLE_COLUMNS.has(`${normalizedTableId}:${normalizedColumnId}`)) {
    return true;
  }

  return AUTO_FIT_FIELD_NAMES.has(String(fieldName || '').trim());
}

export function rectSizeFromPdfBox(rect) {
  if (rect && typeof rect === 'object' && !Array.isArray(rect)) {
    const width = Math.abs(Number(rect.width ?? 0));
    const height = Math.abs(Number(rect.height ?? 0));
    if (width > 0 && height > 0) {
      return { width, height };
    }
    return null;
  }

  if (!Array.isArray(rect) || rect.length < 4) {
    return null;
  }

  const width = Math.abs(Number(rect[2]) - Number(rect[0]));
  const height = Math.abs(Number(rect[3]) - Number(rect[1]));
  if (!(width > 0) || !(height > 0)) {
    return null;
  }

  return { width, height };
}

export function getPdfFieldRectangle(field) {
  try {
    const widgets = field?.acroField?.getWidgets?.();
    if (!Array.isArray(widgets) || widgets.length === 0) {
      return null;
    }

    const rect = widgets[0].getRectangle?.();
    return rectSizeFromPdfBox(rect);
  } catch {
    return null;
  }
}

export function pdfTextContentBox(rect) {
  const size = rectSizeFromPdfBox(rect);
  if (!size) {
    return null;
  }

  const paddingX = Math.min(8, Math.max(2, size.width * 0.04));
  const paddingY = Math.min(6, Math.max(2, size.height * 0.14));
  return {
    width: size.width,
    height: size.height,
    paddingX,
    paddingY,
    maxWidth: Math.max(8, size.width - paddingX * 2),
    maxHeight: Math.max(8, size.height - paddingY * 2)
  };
}

function maxCharsPerLine(maxWidth, fontSize) {
  const charWidth = Math.max(1, fontSize * CHAR_WIDTH_FACTOR);
  return Math.max(4, Math.floor(maxWidth / charWidth));
}

function wrapParagraph(paragraph, maxChars) {
  const compact = String(paragraph || '').replace(/\t/g, ' ').replace(/\s+/g, ' ').trim();
  if (!compact) {
    return [''];
  }

  const words = compact.split(' ').filter(Boolean);
  const lines = [];
  let current = '';

  for (const word of words) {
    const next = current ? `${current} ${word}` : word;
    if (next.length <= maxChars) {
      current = next;
      continue;
    }

    if (current) {
      lines.push(current);
      current = '';
    }

    if (word.length <= maxChars) {
      current = word;
      continue;
    }

    let chunk = '';
    for (const character of word) {
      const candidate = `${chunk}${character}`;
      if (candidate.length > maxChars) {
        if (chunk) {
          lines.push(chunk);
        }
        chunk = character;
      } else {
        chunk = candidate;
      }
    }
    current = chunk;
  }

  if (current) {
    lines.push(current);
  }

  return lines.length > 0 ? lines : [''];
}

export function wrapTextToLines(text, maxWidth, fontSize) {
  const normalized = normalizeMultilineText(text);
  if (!normalized.trim()) {
    return [];
  }

  const maxChars = maxCharsPerLine(maxWidth, fontSize);
  const lines = [];

  for (const paragraph of normalized.split('\n')) {
    lines.push(...wrapParagraph(paragraph, maxChars));
  }

  return lines;
}

function textFitsColumn(lines, fontSize, maxHeight) {
  const lineHeight = fontSize * LINE_HEIGHT_FACTOR;
  return lines.length * lineHeight <= maxHeight;
}

/**
 * Fit multiline PDF column text: measure the available field box, wrap the
 * content, and shrink the font until every character remains visible.
 */
export function prepareMultilinePdfText({
  text,
  rect,
  startFontSize = DEFAULT_START_FONT_SIZE,
  minFontSize = MIN_READABLE_FONT_SIZE
} = {}) {
  const normalized = normalizeMultilineText(text).trim();
  const resolvedStart = Number.isFinite(startFontSize) ? startFontSize : DEFAULT_START_FONT_SIZE;
  const resolvedMin = Number.isFinite(minFontSize) ? minFontSize : MIN_READABLE_FONT_SIZE;

  if (!normalized) {
    return {
      text: '',
      fontSize: resolvedStart,
      lines: [],
      lineHeight: resolvedStart * LINE_HEIGHT_FACTOR
    };
  }

  const box = pdfTextContentBox(rect);
  if (!box) {
    return {
      text: normalized,
      fontSize: resolvedStart,
      lines: normalized.split('\n'),
      lineHeight: resolvedStart * LINE_HEIGHT_FACTOR
    };
  }

  let wrappedLines = wrapTextToLines(normalized, box.maxWidth, resolvedStart);
  for (let fontSize = resolvedStart; fontSize >= resolvedMin; fontSize -= 0.5) {
    wrappedLines = wrapTextToLines(normalized, box.maxWidth, fontSize);
    if (textFitsColumn(wrappedLines, fontSize, box.maxHeight)) {
      return {
        text: wrappedLines.join('\n'),
        fontSize,
        lines: wrappedLines,
        lineHeight: fontSize * LINE_HEIGHT_FACTOR
      };
    }
  }

  wrappedLines = wrapTextToLines(normalized, box.maxWidth, resolvedMin);
  return {
    text: wrappedLines.join('\n'),
    fontSize: resolvedMin,
    lines: wrappedLines,
    lineHeight: resolvedMin * LINE_HEIGHT_FACTOR
  };
}
