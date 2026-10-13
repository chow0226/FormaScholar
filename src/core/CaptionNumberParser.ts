/** Shared caption-number parsing for Arabic and Roman manuscript labels. */
export function romanNumeralToNumber(value: string): number | null {
  const values: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  const normalized = value.trim().toUpperCase();
  // Accept only conventional Roman spellings (1-3999). This prevents
  // malformed labels such as IIII, IC, VX, and IIV from being treated as
  // deterministic numbers during caption validation or auto-fixing.
  const canonicalRoman = /^M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$/;
  if (!normalized || !canonicalRoman.test(normalized)) return null;

  let total = 0;
  for (let index = 0; index < normalized.length; index += 1) {
    const current = values[normalized[index]];
    const next = values[normalized[index + 1]] || 0;
    total += current < next ? -current : current;
  }
  return total > 0 ? total : null;
}

/**
 * Parse the logical number from a FIGURE/Fig. or TABLE caption. This parser
 * intentionally accepts Roman labels for sequence analysis; strict AIP label
 * syntax remains a separate validation rule.
 */
export function parseCaptionNumber(
  text: string,
  kind: 'figure' | 'table'
): number | null {
  const keyword = kind === 'figure' ? '(?:FIGURE|FIG\\.?)' : 'TABLE';
  const match = text.trim().match(
    new RegExp(`^${keyword}\\s+(\\d+[A-Za-z]?|[IVXLCDM]+)`, 'i')
  );
  if (!match) return null;

  const token = match[1];
  if (/^\d/.test(token)) {
    const value = Number(token.match(/^\d+/)?.[0]);
    return Number.isFinite(value) ? value : null;
  }
  return romanNumeralToNumber(token);
}
