const BTB_DATE_SUFFIX = /_(\d{4}-\d{2}-\d{2})$/;

export function sanitizeBtbInput(name: string): string {
  return (
    String(name || '')
      .trim()
      .replace(/\s+/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_|_$/g, '') || 'Baustelle'
  );
}

export function buildBtbTitle(input: string, date?: string): string {
  const sanitized = sanitizeBtbInput(input);
  const day = date || new Date().toISOString().slice(0, 10);
  return `BTB_${sanitized}_${day}`;
}

export function buildFotodokuTitle(input: string, date?: string): string {
  const sanitized = sanitizeBtbInput(input);
  const day = date || new Date().toISOString().slice(0, 10);
  return `BTB_Fotodoku_${sanitized}_${day}`;
}

export function parseBtbTitle(
  title: string
): { input: string; date: string; kind: 'btb' | 'fotodoku' } | null {
  const trimmed = String(title || '').trim();

  const fotodokuMatch = trimmed.match(/^BTB_Fotodoku_(.+)_(\d{4}-\d{2}-\d{2})$/i);
  if (fotodokuMatch) {
    return {
      input: humanizeBtbInput(fotodokuMatch[1]),
      date: fotodokuMatch[2],
      kind: 'fotodoku'
    };
  }

  const btbMatch = trimmed.match(/^BTB_(.+)_(\d{4}-\d{2}-\d{2})$/i);
  if (btbMatch) {
    return {
      input: humanizeBtbInput(btbMatch[1]),
      date: btbMatch[2],
      kind: 'btb'
    };
  }

  const legacyMatch = trimmed.match(/^BTB\s+(\d{4}-\d{2}-\d{2})\s*-\s*(.+)$/i);
  if (legacyMatch) {
    return { input: legacyMatch[2].trim(), date: legacyMatch[1], kind: 'btb' };
  }

  return null;
}

export function fotodokuTitleFromBtbTitle(btbTitle: string): string {
  const parsed = parseBtbTitle(btbTitle);
  if (parsed) {
    return buildFotodokuTitle(parsed.input, parsed.date);
  }
  const withoutDate = btbTitle.replace(BTB_DATE_SUFFIX, '');
  return `BTB_Fotodoku_${sanitizeBtbInput(withoutDate.replace(/^BTB_?/i, ''))}`;
}

function humanizeBtbInput(value: string): string {
  return String(value || '')
    .trim()
    .replace(/_/g, ' ');
}
