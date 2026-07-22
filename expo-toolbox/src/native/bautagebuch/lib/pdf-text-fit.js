const MIN_FONT_SIZE = 6;
const LINE_HEIGHT_FACTOR = 1.22;
const CHAR_WIDTH_FACTOR = 0.52;

export function normalizeMultilineText(text) {
  return String(text ?? '').replace(/\r\n/g, '\n');
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

export function getPdfFieldRectangle(field) {
  try {
    const widgets = field?.acroField?.getWidgets?.();
    if (!Array.isArray(widgets) || widgets.length === 0) {
      return null;
    }

    const rect = widgets[0].getRectangle?.();
    if (!rect) {
      return null;
    }

    const width = Math.abs(Number(rect.width ?? 0));
    const height = Math.abs(Number(rect.height ?? 0));

    if (!(width > 0) || !(height > 0)) {
      return null;
    }

    return { width, height };
  } catch {
    return null;
  }
}

export function prepareMultilinePdfText({
  text,
  rect,
  startFontSize = 12,
  minFontSize = MIN_FONT_SIZE
}) {
  const normalized = normalizeMultilineText(text).trim();
  if (!normalized) {
    return { text: '', fontSize: startFontSize };
  }

  if (!rect) {
    return { text: normalized, fontSize: startFontSize };
  }

  const paddingX = Math.min(8, Math.max(2, rect.width * 0.04));
  const paddingY = Math.min(6, Math.max(2, rect.height * 0.14));
  const maxWidth = Math.max(8, rect.width - paddingX * 2);
  const maxHeight = Math.max(8, rect.height - paddingY * 2);

  for (let fontSize = startFontSize; fontSize >= minFontSize; fontSize -= 0.5) {
    const wrappedLines = wrapTextToLines(normalized, maxWidth, fontSize);
    const lineHeight = fontSize * LINE_HEIGHT_FACTOR;
    const totalHeight = wrappedLines.length * lineHeight;
    if (totalHeight <= maxHeight) {
      return { text: wrappedLines.join('\n'), fontSize };
    }
  }

  const wrappedLines = wrapTextToLines(normalized, maxWidth, minFontSize);
  return { text: wrappedLines.join('\n'), fontSize: minFontSize };
}
