export type FormattingSource = 'direct' | 'style' | 'basedOn' | 'docDefaults' | 'default' | 'unavailable';

export interface LineSpacingValue {
  multiple: number | null;
  rule: string;
  /** Raw w:line value in twentieths of a point. */
  line?: number | null;
}

export interface TabStopDefinition {
  positionPt: number;
  alignment: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'list' | 'unknown';
}

export interface NumberingLevelFormatting {
  format: string | null;
  suffix: string | null;
  /** The separator after the generated number, such as tab or space. */
  separator?: string | null;
  startAt: number | null;
  alignment: string | null;
  numberPositionPt: number | null;
  textPositionPt: number | null;
}

/**
 * Resolve the list-level layer independently from paragraph-style indents.
 * Word stores these values in numbering.xml, not in the Reference paragraph
 * style, so combining them with paragraph formatting would hide which layer
 * actually violates the template.
 */
export function resolveEffectiveNumberingFormatting(
  level: NumberingLevelFormatting | undefined
): NumberingLevelFormatting {
  return {
    format: level?.format ?? null,
    suffix: level?.suffix ?? null,
    separator: level?.separator ?? null,
    startAt: level?.startAt ?? null,
    alignment: level?.alignment ?? null,
    numberPositionPt: level?.numberPositionPt ?? null,
    textPositionPt: level?.textPositionPt ?? null
  };
}

export interface FormattingLayer {
  fontName?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  allCaps?: boolean;
  verticalAlign?: 'superscript' | 'subscript';
  /** OOXML w:position, converted from half-points. */
  baselinePositionPt?: number;
  alignment?: string;
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  pageBreakBefore?: boolean;
  lineSpacing?: LineSpacingValue;
  tabStops?: TabStopDefinition[];
}

export interface StyleFormatting extends FormattingLayer {
  basedOn?: string;
  type?: string;
}

export interface EffectiveRunFormatting {
  text: string;
  fontName: string | null;
  fontSize: number | null;
  bold: boolean | null;
  italic: boolean | null;
  allCaps: boolean | null;
  verticalAlign: 'superscript' | 'subscript' | null;
  baselinePositionPt: number | null;
  sources: Partial<Record<'fontName' | 'fontSize' | 'bold' | 'italic' | 'allCaps' | 'baselinePositionPt', FormattingSource>>;
}

export interface EffectiveParagraphFormatting {
  fontName: string | null;
  fontSize: number | null;
  bold: boolean | null;
  italic: boolean | null;
  allCaps: boolean | null;
  verticalAlign: 'superscript' | 'subscript' | null;
  baselinePositionPt: number | null;
  alignment: string | null;
  firstLineIndent: number | null;
  leftIndent: number | null;
  rightIndent: number | null;
  spaceBefore: number | null;
  spaceAfter: number | null;
  pageBreakBefore: boolean | null;
  lineSpacingMultiple: number | null;
  lineSpacingRule: string | null;
  lineSpacingValue: number | null;
  tabStops: TabStopDefinition[];
  sources: Partial<Record<keyof FormattingLayer | 'lineSpacingMultiple' | 'lineSpacingRule', FormattingSource>>;
}

type ResolvableKey = Exclude<keyof FormattingLayer, 'lineSpacing'>;

const scalarKeys: ResolvableKey[] = [
  'fontName', 'fontSize', 'bold', 'italic', 'allCaps', 'verticalAlign', 'baselinePositionPt', 'alignment',
  'firstLineIndent', 'leftIndent', 'rightIndent', 'spaceBefore', 'spaceAfter', 'pageBreakBefore'
];

/**
 * Resolves paragraph formatting once for the whole validation pipeline.
 * Direct paragraph properties win over the paragraph style, followed by its
 * basedOn chain, docDefaults, and only then the known Word single-spacing
 * default. Missing formatting remains null rather than being invented.
 */
export function resolveEffectiveFormatting(
  direct: FormattingLayer,
  styleId: string | null,
  styles: Map<string, StyleFormatting>,
  defaults: FormattingLayer
): EffectiveParagraphFormatting {
  const sources: EffectiveParagraphFormatting['sources'] = {};

  const resolveScalar = (key: ResolvableKey): string | number | boolean | null => {
    if (direct[key] !== undefined) {
      sources[key] = 'direct';
      return direct[key] as string | number | boolean;
    }

    let current = styleId;
    const seen = new Set<string>();
    let firstStyle = true;
    while (current && !seen.has(current)) {
      seen.add(current);
      const style = styles.get(current);
      if (!style) break;
      if (style[key] !== undefined) {
        sources[key] = firstStyle ? 'style' : 'basedOn';
        return style[key] as string | number | boolean;
      }
      current = style.basedOn || null;
      firstStyle = false;
    }

    if (defaults[key] !== undefined) {
      sources[key] = 'docDefaults';
      return defaults[key] as string | number | boolean;
    }

    // These paragraph/run defaults are defined by Word semantics. Keep them
    // explicit so the validator does not need to fall back to a platform API
    // merely because the DOCX omitted a redundant default element.
    const wordDefault: Partial<Record<ResolvableKey, string | number | boolean>> = {
      bold: false,
      italic: false,
      allCaps: false,
      verticalAlign: undefined,
      alignment: 'left',
      firstLineIndent: 0,
      leftIndent: 0,
      rightIndent: 0,
      spaceBefore: 0,
      spaceAfter: 0,
      pageBreakBefore: false
    };
    if (wordDefault[key] !== undefined) {
      sources[key] = 'default';
      return wordDefault[key] as string | number | boolean;
    }

    sources[key] = 'unavailable';
    return null;
  };

  let lineSpacing: LineSpacingValue | undefined;
  if (direct.lineSpacing !== undefined) {
    lineSpacing = direct.lineSpacing;
    sources.lineSpacingMultiple = 'direct';
    sources.lineSpacingRule = 'direct';
  } else {
    let current = styleId;
    const seen = new Set<string>();
    let firstStyle = true;
    while (current && !seen.has(current)) {
      seen.add(current);
      const style = styles.get(current);
      if (!style) break;
      if (style.lineSpacing !== undefined) {
        lineSpacing = style.lineSpacing;
        sources.lineSpacingMultiple = firstStyle ? 'style' : 'basedOn';
        sources.lineSpacingRule = firstStyle ? 'style' : 'basedOn';
        break;
      }
      current = style.basedOn || null;
      firstStyle = false;
    }
  }

  if (!lineSpacing && defaults.lineSpacing !== undefined) {
    lineSpacing = defaults.lineSpacing;
    sources.lineSpacingMultiple = 'docDefaults';
    sources.lineSpacingRule = 'docDefaults';
  }
  if (!lineSpacing) {
    lineSpacing = { multiple: 1, rule: 'default', line: null };
    sources.lineSpacingMultiple = 'default';
    sources.lineSpacingRule = 'default';
  }

  let tabStops: TabStopDefinition[] | undefined;
  if (direct.tabStops !== undefined) {
    tabStops = direct.tabStops;
  } else {
    let current = styleId;
    const seen = new Set<string>();
    while (current && !seen.has(current)) {
      seen.add(current);
      const style = styles.get(current);
      if (!style) break;
      if (style.tabStops !== undefined) {
        tabStops = style.tabStops;
        break;
      }
      current = style.basedOn || null;
    }
  }
  if (tabStops === undefined) tabStops = defaults.tabStops ?? [];

  const values = Object.fromEntries(scalarKeys.map(key => [key, resolveScalar(key)])) as Record<ResolvableKey, string | number | boolean | null>;
  return {
    fontName: values.fontName as string | null,
    fontSize: values.fontSize as number | null,
    bold: values.bold as boolean | null,
    italic: values.italic as boolean | null,
    allCaps: values.allCaps as boolean | null,
    verticalAlign: values.verticalAlign as 'superscript' | 'subscript' | null,
    baselinePositionPt: values.baselinePositionPt as number | null,
    alignment: values.alignment as string | null,
    firstLineIndent: values.firstLineIndent as number | null,
    leftIndent: values.leftIndent as number | null,
    rightIndent: values.rightIndent as number | null,
    spaceBefore: values.spaceBefore as number | null,
    spaceAfter: values.spaceAfter as number | null,
    pageBreakBefore: values.pageBreakBefore as boolean | null,
    lineSpacingMultiple: lineSpacing.multiple,
    lineSpacingRule: lineSpacing.rule,
    lineSpacingValue: lineSpacing.line ?? null,
    tabStops,
    sources
  };
}

/** Resolve a text run after paragraph-level formatting has been resolved. */
export function resolveEffectiveRunFormatting(
  text: string,
  direct: FormattingLayer,
  characterStyleId: string | null,
  paragraph: EffectiveParagraphFormatting,
  styles: Map<string, StyleFormatting>,
  defaults: FormattingLayer
): EffectiveRunFormatting {
  const keys: Array<'fontName' | 'fontSize' | 'bold' | 'italic' | 'allCaps' | 'baselinePositionPt'> = [
    'fontName', 'fontSize', 'bold', 'italic', 'allCaps', 'baselinePositionPt'
  ];
  const sources: EffectiveRunFormatting['sources'] = {};

  const resolve = (key: typeof keys[number]): string | number | boolean | null => {
    if (direct[key] !== undefined) {
      sources[key] = 'direct';
      return direct[key] as string | number | boolean;
    }

    let current = characterStyleId;
    const seen = new Set<string>();
    let firstStyle = true;
    while (current && !seen.has(current)) {
      seen.add(current);
      const style = styles.get(current);
      if (!style) break;
      if (style[key] !== undefined) {
        sources[key] = firstStyle ? 'style' : 'basedOn';
        return style[key] as string | number | boolean;
      }
      current = style.basedOn || null;
      firstStyle = false;
    }

    const paragraphValue = paragraph[key];
    if (paragraphValue !== null && paragraphValue !== undefined) {
      sources[key] = paragraph.sources[key] || 'style';
      return paragraphValue as string | number | boolean;
    }

    if (defaults[key] !== undefined) {
      sources[key] = 'docDefaults';
      return defaults[key] as string | number | boolean;
    }

    sources[key] = 'unavailable';
    return null;
  };

  const values = Object.fromEntries(keys.map(key => [key, resolve(key)])) as Record<typeof keys[number], string | number | boolean | null>;
  const resolveVerticalAlign = (): 'superscript' | 'subscript' | null => {
    if (direct.verticalAlign !== undefined) return direct.verticalAlign;

    let current = characterStyleId;
    const seen = new Set<string>();
    while (current && !seen.has(current)) {
      seen.add(current);
      const style = styles.get(current);
      if (!style) break;
      if (style.verticalAlign !== undefined) return style.verticalAlign;
      current = style.basedOn || null;
    }

    if (paragraph.verticalAlign !== null) return paragraph.verticalAlign;
    return defaults.verticalAlign || null;
  };
  return {
    text,
    fontName: values.fontName as string | null,
    fontSize: values.fontSize as number | null,
    bold: values.bold as boolean | null,
    italic: values.italic as boolean | null,
    allCaps: values.allCaps as boolean | null,
    verticalAlign: resolveVerticalAlign(),
    baselinePositionPt: values.baselinePositionPt as number | null,
    sources
  };
}
