export const approxEqual = (a: number | null | undefined, b: number, tolerance = 0.6): boolean =>
  typeof a === 'number' && Math.abs(a - b) <= tolerance;

/** Stable text comparison form used by Office/OOXML and stale-fix checks. */
export const paragraphTextFingerprint = (value: string): string => value
  .normalize('NFKC')
  .replace(/\u00A0/g, ' ')
  .replace(/[\r\n\v\f]+/g, '\n')
  .replace(/\s+/g, ' ')
  .trim();

export const displayValue = (value: unknown, unit?: string): string => {
  if (value === null || value === undefined || value === '') return 'Not detected';
  if (typeof value === 'boolean') return value ? 'Yes' : 'No';
  if (typeof value === 'number') {
    const rounded = Math.round(value * 100) / 100;
    return unit ? `${rounded} ${unit}` : `${rounded}`;
  }
  return String(value);
};

export const normalizeAlignment = (value: string): string => {
  const v = (value || '').toLowerCase();
  if (v === 'both') return 'justified';
  if (v === 'center' || v === 'centered') return 'centered';
  if (v === 'left') return 'left';
  if (v === 'right') return 'right';
  if (v === 'justified') return 'justified';
  return v;
};

export const isAllCaps = (text: string): boolean => {
  const letters = text.replace(/[^A-Za-z]/g, '');
  return letters.length > 0 && letters === letters.toUpperCase();
};

const minorWords = new Set(['a', 'an', 'the', 'and', 'but', 'or', 'nor', 'for', 'so', 'yet', 'at', 'by', 'in', 'of', 'off', 'on', 'per', 'to', 'up', 'via', 'as', 'from', 'with']);

/**
 * Technical terms may intentionally use mixed case, digits, or a known
 * technical hyphenated spelling. They must not be normalized like ordinary
 * title words.
 */
const PRESERVED_TECHNICAL_TOKENS = new Set([
  'mRMR',
  'XGBoost',
  'eHealth',
  'scikit-learn',
  'iOS',
  'pH'
]);

export const isTechnicalToken = (word: string): boolean => {
  const clean = word.replace(/^[^A-Za-z0-9]+|[^A-Za-z0-9-]+$/g, '');
  if (!clean) return false;
  if (PRESERVED_TECHNICAL_TOKENS.has(clean)) return true;
  if (/\d/.test(clean)) return true;
  return false;
};

export const looksTitleCase = (text: string): boolean => {
  const words = text.trim().replace(/[()]/g, '').split(/\s+/).filter(Boolean);
  if (words.length === 0) return false;
  return words.every((word, index) => {
    const clean = word.replace(/^[^A-Za-z]+|[^A-Za-z]+$/g, '');
    if (!clean) return true;
    if (isTechnicalToken(clean)) return true;
    const lower = clean.toLowerCase();
    if (index > 0 && index < words.length - 1 && minorWords.has(lower)) return clean === lower;
    return clean[0] === clean[0].toUpperCase();
  });
};

/**
 * Checks only the significant words in a heading. Function words are handled
 * by the separate minor-word rule so one capitalization choice produces one
 * finding instead of two overlapping failures.
 */
export const looksMajorTitleCase = (text: string): boolean => {
  const words = text.trim().replace(/[()]/g, '').split(/\s+/).filter(Boolean);
  if (words.length === 0) return false;
  const minor = new Set(['a', 'an', 'the', 'and', 'but', 'or', 'nor', 'for', 'from', 'with', 'in', 'on', 'at', 'by', 'to', 'of', 'off', 'as', 'via', 'vs', 'versus', 'per']);
  return words.every((word, index) => {
    const clean = word.replace(/^[^A-Za-z]+|[^A-Za-z]+$/g, '');
    if (!clean) return true;
    if (isTechnicalToken(clean)) return true;
    if (index > 0 && index < words.length - 1 && minor.has(clean.toLowerCase())) return true;
    return clean[0] === clean[0].toUpperCase();
  });
};

export const isBlank = (text: string): boolean => !text || text.trim().length === 0;
