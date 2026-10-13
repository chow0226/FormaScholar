import JSZip from 'jszip';
import { RuleDefinition, StrategyDefinition } from '../core/models';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export type TemplateConfidence = 'high' | 'medium' | 'manual';
export type TemplateProposalStatus = 'unchanged' | 'changed' | 'new' | 'not-determinable';

export interface TemplateEvidence {
  source: 'style' | 'repeated-examples' | 'document-structure';
  styleName?: string;
  paragraphCount?: number;
  note: string;
}

export interface TemplateRuleCandidate {
  ruleId?: string;
  component: string;
  property: string;
  detectedValue: unknown;
  confidence: TemplateConfidence;
  evidence: TemplateEvidence;
}

export interface StrategyChangeProposal {
  ruleId: string;
  ruleName: string;
  component: string;
  property: string;
  currentValue?: unknown;
  detectedValue?: unknown;
  status: TemplateProposalStatus;
  confidence: TemplateConfidence;
  evidence?: TemplateEvidence;
  note?: string;
}

export interface TemplateComponentSummary {
  component: string;
  styleName?: string;
  examples: number;
  confidence: TemplateConfidence;
}

export interface TemplateAnalysisResult {
  fileName: string;
  components: TemplateComponentSummary[];
  candidates: TemplateRuleCandidate[];
  proposals: StrategyChangeProposal[];
  counts: {
    unchanged: number;
    changed: number;
    new: number;
    review: number;
    notDeterminable: number;
  };
  warnings: string[];
}

const SHARED_SPACING_TOLERANCE_PT = 0.05;
const FLOATING_POINT_EPSILON = 1e-9;

function numbersNearEqual(
  left: number,
  right: number,
  tolerance = SHARED_SPACING_TOLERANCE_PT
): boolean {
  return Math.abs(left - right) <= tolerance + FLOATING_POINT_EPSILON;
}

/** Resolve a nearly-equal before/after pair to the single value supported by the strategy schema. */
export function sharedSpacingValue(value: unknown): number | undefined {
  if (!value || typeof value !== 'object' || !('before' in value) || !('after' in value)) {
    return undefined;
  }

  const pair = value as { before: unknown; after: unknown };
  if (typeof pair.before !== 'number'
    || typeof pair.after !== 'number'
    || !Number.isFinite(pair.before)
    || !Number.isFinite(pair.after)
    || !numbersNearEqual(pair.before, pair.after)) {
    return undefined;
  }

  return Math.round(((pair.before + pair.after) / 2) * 20) / 20;
}

export function isTemplateProposalAcceptable(proposal: StrategyChangeProposal): boolean {
  if (proposal.status !== 'changed' || proposal.confidence === 'manual') return false;
  const detected = proposal.detectedValue;
  if (sharedSpacingValue(detected) !== undefined) return true;
  return typeof detected === 'string' || typeof detected === 'number' || typeof detected === 'boolean';
}

/** Return only proposal IDs that the current strategy-copy workflow can apply safely. */
export function safeTemplateProposalIds(analysis: TemplateAnalysisResult): string[] {
  return analysis.proposals
    .filter(isTemplateProposalAcceptable)
    .map(proposal => proposal.ruleId);
}

interface Formatting {
  fontName?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  alignment?: string;
  spaceBefore?: number;
  spaceAfter?: number;
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  verticalAlign?: string;
  baselinePositionPt?: number;
}

interface TemplateStyle extends Formatting {
  id: string;
  name: string;
  basedOn?: string;
}

interface ParsedTemplate {
  styles: Map<string, TemplateStyle>;
  paragraphs: Array<{ text: string; styleId?: string; formatting: Formatting; runFormattings: Formatting[] }>;
  layout: Record<string, number | boolean | undefined>;
}

function wAttr(element: Element | null | undefined, name: string): string | undefined {
  if (!element) return undefined;
  return element.getAttributeNS(W_NS, name)
    || element.getAttribute(`w:${name}`)
    || element.getAttribute(name)
    || undefined;
}

function child(element: Element | null | undefined, localName: string): Element | undefined {
  if (!element) return undefined;
  return Array.from(element.children).find(item => item.localName === localName);
}

function numberValue(value: string | undefined, divisor = 1): number | undefined {
  if (value === undefined) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed / divisor : undefined;
}

function normalizeAlignment(value: string | undefined): string | undefined {
  if (!value) return undefined;
  switch (value.toLowerCase()) {
    case 'both': return 'justified';
    case 'center': return 'centered';
    case 'start': return 'left';
    case 'end': return 'right';
    default: return value.toLowerCase();
  }
}

function formatTemplateValue(value: unknown, unit?: string): string {
  if (value && typeof value === 'object' && 'before' in value && 'after' in value) {
    const pair = value as { before: unknown; after: unknown };
    return `${String(pair.before)} / ${String(pair.after)} ${unit || 'pt'}`;
  }
  if (typeof value === 'boolean') return value ? 'Yes' : 'No';
  if (typeof value === 'number') return `${String(value)}${unit ? ` ${unit}` : ''}`;
  return String(value);
}

function reviewedTemplateEvidence(
  proposal: StrategyChangeProposal,
  fileName: string,
  detected: unknown,
  unit?: string
): string {
  const note = proposal.evidence?.note?.trim().replace(/[.]+$/, '');
  const detail = note ? `${note}. ` : '';
  return `Reviewed template ${fileName}: ${detail}Detected value: ${formatTemplateValue(detected, unit)}.`;
}

function booleanValue(element: Element | undefined): boolean | undefined {
  if (!element) return undefined;
  const value = wAttr(element, 'val');
  if (value === undefined) return true;
  return !['0', 'false', 'off', 'no'].includes(value.toLowerCase());
}

function parseFormatting(parent: Element | null | undefined): Formatting {
  const pPr = child(parent, 'pPr');
  const rPr = child(parent, 'rPr');
  const rFonts = child(rPr, 'rFonts');
  const spacing = child(pPr, 'spacing');
  const ind = child(pPr, 'ind');
  const jc = child(pPr, 'jc');

  return {
    fontName: wAttr(rFonts, 'ascii') || wAttr(rFonts, 'hAnsi') || wAttr(rFonts, 'eastAsia'),
    fontSize: numberValue(wAttr(child(rPr, 'sz'), 'val'), 2),
    bold: booleanValue(child(rPr, 'b')),
    italic: booleanValue(child(rPr, 'i')),
    alignment: normalizeAlignment(wAttr(jc, 'val')),
    spaceBefore: numberValue(wAttr(spacing, 'before'), 20),
    spaceAfter: numberValue(wAttr(spacing, 'after'), 20),
    firstLineIndent: numberValue(wAttr(ind, 'firstLine'), 20),
    leftIndent: numberValue(wAttr(ind, 'left'), 20),
    rightIndent: numberValue(wAttr(ind, 'right'), 20),
    verticalAlign: wAttr(child(rPr, 'vertAlign'), 'val')?.toLowerCase(),
    baselinePositionPt: numberValue(wAttr(child(rPr, 'position'), 'val'), 2)
  };
}

function mergeFormatting(base: Formatting, direct: Formatting): Formatting {
  return {
    ...base,
    ...Object.fromEntries(Object.entries(direct).filter(([, value]) => value !== undefined))
  };
}

function textOf(element: Element): string {
  return Array.from(element.getElementsByTagNameNS(W_NS, 't'))
    .map(node => node.textContent || '')
    .join('');
}

function normalizeComponentText(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function componentForStyle(style: TemplateStyle): string | undefined {
  const value = normalizeComponentText(`${style.name} ${style.id}`);
  if (/paper title|^title\b/.test(value)) return 'title';
  if (/author.*name|name.*author/.test(value)) return 'authorName';
  if (/affiliation/.test(value)) return 'authorAffiliation';
  if (/e ?mail|email/.test(value)) return 'authorEmail';
  if (/abstract/.test(value)) return 'abstract';
  if (/heading ?1|^h1$/.test(value)) return 'heading1';
  if (/heading ?2|^h2$/.test(value)) return 'heading2';
  if (/heading ?3|^h3$/.test(value)) return 'heading3';
  if (/figure.*caption|caption.*figure/.test(value)) return 'figureCaption';
  if (/table.*caption|caption.*table/.test(value)) return 'tableCaption';
  if (/reference/.test(value)) return 'referencesHeading';
  return undefined;
}

function componentForParagraph(text: string, style?: TemplateStyle): string | undefined {
  const fromStyle = style && componentForStyle(style);
  if (fromStyle) return fromStyle;
  const normalized = text.trim().toLowerCase();
  if (/^abstract\./.test(normalized)) return 'abstract';
  if (/^references?$/.test(normalized)) return 'referencesHeading';
  if (/^table\s+(?:\d+|[ivxlcdm]+)\s*[.:]/i.test(text)) return 'tableCaption';
  if (/^figure\s+(?:\d+|[ivxlcdm]+)\s*[.:]/i.test(text)) return 'figureCaption';
  return undefined;
}

const RUN_FORMATTING_KEYS: Array<keyof Formatting> = ['fontName', 'fontSize', 'bold', 'italic'];

function isRaisedFormatting(formatting: Formatting): boolean {
  return formatting.verticalAlign === 'superscript'
    || (formatting.baselinePositionPt ?? 0) > 0;
}

function consensusRunFormatting(base: Formatting, runs: Formatting[]): Formatting {
  if (!runs.length) return base;
  const result: Formatting = { ...base };
  const ordinaryRuns = runs.filter(run => !isRaisedFormatting(run));
  for (const key of RUN_FORMATTING_KEYS) {
    const values = ordinaryRuns
      .map(run => run[key])
      .filter((value): value is string | number | boolean => value !== undefined);
    if (!values.length) continue;
    if (values.every(value => valuesEqual(value, values[0]))) {
      (result as unknown as Record<keyof Formatting, unknown>)[key] = values[0];
    }
    else delete result[key];
  }
  return result;
}

function parseTemplateWithDom(documentXml: string, stylesXml: string): ParsedTemplate {
  const parser = new DOMParser();
  const stylesDoc = parser.parseFromString(stylesXml, 'application/xml');
  const documentDoc = parser.parseFromString(documentXml, 'application/xml');
  const styles = new Map<string, TemplateStyle>();

  for (const element of Array.from(stylesDoc.getElementsByTagNameNS(W_NS, 'style'))) {
    const id = wAttr(element, 'styleId');
    if (!id) continue;
    const name = wAttr(child(element, 'name'), 'val') || id;
    styles.set(id, {
      id,
      name,
      basedOn: wAttr(child(element, 'basedOn'), 'val'),
      ...parseFormatting(element)
    });
  }

  const resolveStyle = (id: string, seen = new Set<string>()): TemplateStyle | undefined => {
    const current = styles.get(id);
    if (!current || seen.has(id)) return current;
    const nextSeen = new Set(seen).add(id);
    const base = current.basedOn ? resolveStyle(current.basedOn, nextSeen) : undefined;
    if (!base) return current;
    return { ...base, ...current, ...mergeFormatting(base, current) };
  };

  const paragraphs: ParsedTemplate['paragraphs'] = [];
  for (const paragraph of Array.from(documentDoc.getElementsByTagNameNS(W_NS, 'p'))) {
    const pPr = child(paragraph, 'pPr');
    const styleId = wAttr(child(pPr, 'pStyle'), 'val');
    const style = styleId ? resolveStyle(styleId) : undefined;
    const paragraphFormatting = mergeFormatting(style || {}, parseFormatting(paragraph));
    const runFormattings = Array.from(paragraph.getElementsByTagNameNS(W_NS, 'r'))
      .filter(run => textOf(run).trim().length > 0)
      .map(run => mergeFormatting(paragraphFormatting, parseFormatting(run)));
    paragraphs.push({
      text: textOf(paragraph),
      styleId,
      formatting: consensusRunFormatting(paragraphFormatting, runFormattings),
      runFormattings
    });
  }

  const sections = Array.from(documentDoc.getElementsByTagNameNS(W_NS, 'sectPr'));
  const section = sections[sections.length - 1];
  const pageSize = child(section, 'pgSz');
  const margins = child(section, 'pgMar');
  const columns = child(section, 'cols');
  const twipsToInches = (value: string | undefined): number | undefined => numberValue(value, 1440);
  const layout: ParsedTemplate['layout'] = {
    pageWidthInches: twipsToInches(wAttr(pageSize, 'w')),
    pageHeightInches: twipsToInches(wAttr(pageSize, 'h')),
    marginTopInches: twipsToInches(wAttr(margins, 'top')),
    marginRightInches: twipsToInches(wAttr(margins, 'right')),
    marginBottomInches: twipsToInches(wAttr(margins, 'bottom')),
    marginLeftInches: twipsToInches(wAttr(margins, 'left')),
    columns: Number(wAttr(columns, 'num') || '1') || 1
  };

  const resolvedStyles = new Map<string, TemplateStyle>();
  for (const id of styles.keys()) {
    const resolved = resolveStyle(id);
    if (resolved) resolvedStyles.set(id, resolved);
  }
  return { styles: resolvedStyles, paragraphs, layout };
}

function xmlAttribute(xml: string, name: string): string | undefined {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return xml.match(new RegExp(`(?:^|\\s)(?:[A-Za-z_][\\w.-]*:)?${escaped}\\s*=\\s*["']([^"']*)["']`, 'i'))?.[1];
}

function xmlBlock(xml: string, localName: string): string | undefined {
  const escaped = localName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return xml.match(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${escaped}\\b[^>]*>[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?${escaped}\\s*>`, 'i'))?.[0];
}

function xmlBlocks(xml: string, localName: string): string[] {
  const escaped = localName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return Array.from(xml.matchAll(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${escaped}\\b[^>]*>[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?${escaped}\\s*>`, 'gi')), match => match[0]);
}

function xmlStartTag(xml: string, localName: string): string | undefined {
  const escaped = localName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return xml.match(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${escaped}\\b[^>]*>`, 'i'))?.[0];
}

function xmlValue(xml: string, localName: string, attribute = 'val'): string | undefined {
  const block = xmlBlock(xml, localName) || xmlStartTag(xml, localName);
  return block ? xmlAttribute(block, attribute) : undefined;
}

function booleanXmlValue(xml: string, localName: string): boolean | undefined {
  const tag = xmlStartTag(xmlBlock(xml, localName) || xml, localName);
  if (!tag) return undefined;
  const value = xmlAttribute(tag, 'val');
  return value === undefined || !['0', 'false', 'off', 'no'].includes(value.toLowerCase());
}

function parseCharacterFormattingFromXml(rPr: string): Formatting {
  const rFonts = xmlStartTag(rPr, 'rFonts') || '';
  return {
    fontName: xmlAttribute(rFonts, 'ascii') || xmlAttribute(rFonts, 'hAnsi') || xmlAttribute(rFonts, 'eastAsia'),
    fontSize: numberValue(xmlValue(rPr, 'sz'), 2),
    bold: booleanXmlValue(rPr, 'b'),
    italic: booleanXmlValue(rPr, 'i'),
    verticalAlign: xmlValue(rPr, 'vertAlign')?.toLowerCase(),
    baselinePositionPt: numberValue(xmlValue(rPr, 'position'), 2)
  };
}

function parseParagraphFormattingFromXml(paragraphXml: string): Formatting {
  const pPr = xmlBlock(paragraphXml, 'pPr') || '';
  const paragraphMarkRPr = xmlBlock(pPr, 'rPr') || '';
  const spacing = xmlStartTag(pPr, 'spacing') || '';
  const ind = xmlStartTag(pPr, 'ind') || '';
  const jc = xmlStartTag(pPr, 'jc') || '';
  return {
    ...parseCharacterFormattingFromXml(paragraphMarkRPr),
    alignment: normalizeAlignment(xmlAttribute(jc, 'val')),
    spaceBefore: numberValue(xmlAttribute(spacing, 'before'), 20),
    spaceAfter: numberValue(xmlAttribute(spacing, 'after'), 20),
    firstLineIndent: numberValue(xmlAttribute(ind, 'firstLine'), 20),
    leftIndent: numberValue(xmlAttribute(ind, 'left'), 20),
    rightIndent: numberValue(xmlAttribute(ind, 'right'), 20)
  };
}

function parseRunFormattingFromXml(runXml: string): Formatting {
  return parseCharacterFormattingFromXml(xmlBlock(runXml, 'rPr') || '');
}

function parseFormattingFromXml(parentXml: string): Formatting {
  const pPr = parseParagraphFormattingFromXml(parentXml);
  const rPr = xmlBlock(parentXml, 'rPr') || '';
  return mergeFormatting(pPr, parseCharacterFormattingFromXml(rPr));
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function textFromXml(xml: string): string {
  return Array.from(xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t\s*>/gi), match => decodeXmlText(match[1]))
    .join('');
}

function parseTemplateWithRegex(documentXml: string, stylesXml: string): ParsedTemplate {
  const styles = new Map<string, TemplateStyle>();
  for (const styleXml of xmlBlocks(stylesXml, 'style')) {
    const id = xmlAttribute(xmlStartTag(styleXml, 'style') || '', 'styleId');
    if (!id) continue;
    styles.set(id, {
      id,
      name: xmlValue(styleXml, 'name') || id,
      basedOn: xmlValue(styleXml, 'basedOn'),
      ...parseFormattingFromXml(styleXml)
    });
  }
  const resolveStyle = (id: string, seen = new Set<string>()): TemplateStyle | undefined => {
    const current = styles.get(id);
    if (!current || seen.has(id)) return current;
    const base = current.basedOn ? resolveStyle(current.basedOn, new Set(seen).add(id)) : undefined;
    return base ? { ...base, ...current, ...mergeFormatting(base, current) } : current;
  };

  const paragraphs = xmlBlocks(documentXml, 'p').map(paragraphXml => {
    const pPr = xmlBlock(paragraphXml, 'pPr') || '';
    const styleId = xmlValue(pPr, 'pStyle');
    const style = styleId ? resolveStyle(styleId) : undefined;
    const paragraphFormatting = mergeFormatting(style || {}, parseParagraphFormattingFromXml(paragraphXml));
    const runFormattings = xmlBlocks(paragraphXml, 'r')
      .filter(runXml => textFromXml(runXml).trim().length > 0)
      .map(runXml => mergeFormatting(paragraphFormatting, parseRunFormattingFromXml(runXml)));
    return {
      text: textFromXml(paragraphXml),
      styleId,
      formatting: consensusRunFormatting(paragraphFormatting, runFormattings),
      runFormattings
    };
  });

  const sections = xmlBlocks(documentXml, 'sectPr');
  const section = sections[sections.length - 1] || '';
  // pgSz, pgMar and cols are normally self-closing OOXML elements. Read
  // their start tags directly so the non-browser fallback preserves columns
  // and non-default page geometry too.
  const pageSize = xmlStartTag(section, 'pgSz') || '';
  const margins = xmlStartTag(section, 'pgMar') || '';
  const columns = xmlStartTag(section, 'cols') || '';
  const twipsToInches = (value: string | undefined): number | undefined => numberValue(value, 1440);
  const resolvedStyles = new Map<string, TemplateStyle>();
  for (const id of styles.keys()) {
    const resolved = resolveStyle(id);
    if (resolved) resolvedStyles.set(id, resolved);
  }
  return {
    styles: resolvedStyles,
    paragraphs,
    layout: {
      pageWidthInches: twipsToInches(xmlAttribute(pageSize, 'w')),
      pageHeightInches: twipsToInches(xmlAttribute(pageSize, 'h')),
      marginTopInches: twipsToInches(xmlAttribute(margins, 'top')),
      marginRightInches: twipsToInches(xmlAttribute(margins, 'right')),
      marginBottomInches: twipsToInches(xmlAttribute(margins, 'bottom')),
      marginLeftInches: twipsToInches(xmlAttribute(margins, 'left')),
      columns: Number(xmlAttribute(columns, 'num') || '1') || 1
    }
  };
}

function parseTemplate(documentXml: string, stylesXml: string): ParsedTemplate {
  return typeof DOMParser === 'undefined'
    ? parseTemplateWithRegex(documentXml, stylesXml)
    : parseTemplateWithDom(documentXml, stylesXml);
}

function ruleComponent(rule: RuleDefinition): string {
  return rule.target;
}

function valuesEqual(left: unknown, right: unknown): boolean {
  if (typeof left === 'number' && typeof right === 'number') return numbersNearEqual(left, right);
  if (typeof left === 'string' && typeof right === 'string') return left.trim().toLowerCase() === right.trim().toLowerCase();
  if (typeof left === 'boolean' && typeof right === 'boolean') return left === right;
  if (left && right && typeof left === 'object' && typeof right === 'object') {
    const a = left as Record<string, unknown>;
    const b = right as Record<string, unknown>;
    return valuesEqual(a.before, b.before) && valuesEqual(a.after, b.after);
  }
  return left === right;
}

function currentValueFor(rule: RuleDefinition): unknown {
  if (rule.property === 'paragraph.frontMatterSpacing') {
    return { before: rule.expected, after: rule.expected };
  }
  return rule.expected;
}

function formattingValue(formatting: Formatting, property: string): unknown {
  switch (property) {
    case 'font.name': return formatting.fontName;
    case 'font.size': return formatting.fontSize;
    case 'font.bold': return formatting.bold;
    case 'font.italic': return formatting.italic;
    case 'paragraph.alignment': return formatting.alignment;
    case 'paragraph.frontMatterSpacing':
      return formatting.spaceBefore !== undefined && formatting.spaceAfter !== undefined
        ? { before: formatting.spaceBefore, after: formatting.spaceAfter }
        : undefined;
    case 'paragraph.spaceBefore': return formatting.spaceBefore;
    case 'paragraph.spaceAfter': return formatting.spaceAfter;
    case 'paragraph.firstLineIndent': return formatting.firstLineIndent;
    case 'paragraph.leftIndent': return formatting.leftIndent;
    case 'paragraph.rightIndent': return formatting.rightIndent;
    default: return undefined;
  }
}

function displayComponent(component: string): string {
  return component.replace(/([a-z])([A-Z])/g, '$1 $2');
}

function candidateKey(component: string, property: string): string {
  return `${component}::${property}`;
}

function exactComponentStyleName(component: string, style: TemplateStyle): boolean {
  const name = normalizeComponentText(style.name).replace(/\s+/g, '');
  const exact: Record<string, string[]> = {
    title: ['title', 'papertitle'],
    authorName: ['authorname'],
    authorAffiliation: ['authoraffiliation', 'affiliation'],
    authorEmail: ['authoremail', 'email'],
    abstract: ['abstract'],
    heading1: ['heading1', 'h1'],
    heading2: ['heading2', 'h2'],
    heading3: ['heading3', 'h3'],
    figureCaption: ['figurecaption'],
    tableCaption: ['tablecaption'],
    referencesHeading: ['references', 'bibliography']
  };
  return exact[component]?.includes(name) || false;
}

function styleFormattingSignature(style: TemplateStyle): string {
  return JSON.stringify({
    fontName: style.fontName,
    fontSize: style.fontSize,
    bold: style.bold,
    italic: style.italic,
    alignment: style.alignment,
    spaceBefore: style.spaceBefore,
    spaceAfter: style.spaceAfter,
    firstLineIndent: style.firstLineIndent,
    leftIndent: style.leftIndent,
    rightIndent: style.rightIndent
  });
}

export class TemplateAnalyzer {
  static async analyzeFile(file: File, strategy: StrategyDefinition): Promise<TemplateAnalysisResult> {
    return this.analyzeBytes(await file.arrayBuffer(), file.name, strategy);
  }

  static async analyzeBytes(bytes: ArrayBuffer | Uint8Array, fileName: string, strategy: StrategyDefinition): Promise<TemplateAnalysisResult> {
    const zip = await JSZip.loadAsync(bytes);
    const documentPart = zip.file('word/document.xml');
    if (!documentPart) throw new Error('The uploaded file is not a DOCX template: word/document.xml is missing.');
    const stylesPart = zip.file('word/styles.xml');
    const parsed = parseTemplate(await documentPart.async('string'), stylesPart ? await stylesPart.async('string') : '');
    const warnings: string[] = [];
    if (!stylesPart) warnings.push('word/styles.xml is missing; only repeated document examples could be inspected.');
    if (!parsed.paragraphs.length) warnings.push('No body paragraphs were found in the uploaded template.');

    const componentExamples = new Map<string, Array<{ formatting: Formatting; style?: TemplateStyle }>>();
    for (const paragraph of parsed.paragraphs) {
      const style = paragraph.styleId ? parsed.styles.get(paragraph.styleId) : undefined;
      const component = componentForParagraph(paragraph.text, style);
      if (!component || !paragraph.text.trim()) continue;
      const list = componentExamples.get(component) || [];
      list.push({ formatting: paragraph.formatting, style });
      componentExamples.set(component, list);
    }

    const stylesByComponent = new Map<string, TemplateStyle[]>();
    for (const style of parsed.styles.values()) {
      const component = componentForStyle(style);
      if (!component) continue;
      const list = stylesByComponent.get(component) || [];
      list.push(style);
      stylesByComponent.set(component, list);
    }
    const usedStyleCounts = new Map<string, Map<string, number>>();
    for (const paragraph of parsed.paragraphs) {
      const style = paragraph.styleId ? parsed.styles.get(paragraph.styleId) : undefined;
      const component = componentForParagraph(paragraph.text, style);
      if (!component || !paragraph.styleId) continue;
      const counts = usedStyleCounts.get(component) || new Map<string, number>();
      counts.set(paragraph.styleId, (counts.get(paragraph.styleId) || 0) + 1);
      usedStyleCounts.set(component, counts);
    }

    const componentStyles = new Map<string, TemplateStyle>();
    for (const [component, styles] of stylesByComponent) {
      const usage = usedStyleCounts.get(component);
      const usedStyles = usage
        ? styles.filter(style => usage.has(style.id)).sort((a, b) => (usage.get(b.id) || 0) - (usage.get(a.id) || 0))
        : [];
      if (usedStyles.length > 1) {
        const signatures = new Set(usedStyles.map(styleFormattingSignature));
        if (signatures.size > 1) {
          warnings.push(`${displayComponent(component)} has multiple used styles with conflicting formatting; style-defined proposals require manual review.`);
          continue;
        }
      }
      const exactStyle = styles.find(style => exactComponentStyleName(component, style));
      componentStyles.set(component, usedStyles[0] || exactStyle || styles[0]);
    }

    const candidates: TemplateRuleCandidate[] = [];
    const candidateByKey = new Map<string, TemplateRuleCandidate>();
    const addCandidate = (candidate: TemplateRuleCandidate): void => {
      if (candidate.detectedValue === undefined) return;
      const key = candidateKey(candidate.component, candidate.property);
      const previous = candidateByKey.get(key);
      if (!previous || (previous.confidence !== 'high' && candidate.confidence === 'high')) {
        candidateByKey.set(key, candidate);
      }
    };

    for (const rule of strategy.rules) {
      if (rule.target === 'document') {
        const property = rule.property.replace(/^layout\./, '');
        const detectedValue = parsed.layout[property];
        if (detectedValue !== undefined) addCandidate({
          ruleId: rule.id,
          component: 'document',
          property: rule.property,
          detectedValue,
          confidence: 'high',
          evidence: { source: 'document-structure', note: 'Extracted from word/document.xml section properties.' }
        });
        continue;
      }

      const style = componentStyles.get(ruleComponent(rule));
      const examples = componentExamples.get(ruleComponent(rule)) || [];
      const styleValue = style ? formattingValue(style, rule.property) : undefined;
      const values = examples
        .map(example => formattingValue(example.formatting, rule.property))
        .filter(value => value !== undefined);
      const examplesConsistent = values.length >= 2
        && values.every(value => valuesEqual(value, values[0]));

      if (styleValue !== undefined && examplesConsistent && !valuesEqual(styleValue, values[0])) {
        warnings.push(
          `${displayComponent(ruleComponent(rule))} ${rule.property} uses ${String(styleValue)} in the named style but ${String(values[0])} across repeated template examples.`
        );
        addCandidate({
          ruleId: rule.id,
          component: ruleComponent(rule),
          property: rule.property,
          detectedValue: values[0],
          confidence: 'manual',
          evidence: {
            source: 'repeated-examples',
            styleName: style?.name,
            paragraphCount: values.length,
            note: 'Named style and repeated direct formatting disagree. Manual confirmation is required.'
          }
        });
        continue;
      }

      if (styleValue !== undefined) {
        addCandidate({
          ruleId: rule.id,
          component: ruleComponent(rule),
          property: rule.property,
          detectedValue: styleValue,
          confidence: 'high',
          evidence: { source: 'style', styleName: style?.name, note: 'Taken from the named DOCX style definition.' }
        });
        continue;
      }

      if (!values.length) continue;
      const first = values[0];
      if (values.every(value => valuesEqual(value, first))) {
        addCandidate({
          ruleId: rule.id,
          component: ruleComponent(rule),
          property: rule.property,
          detectedValue: first,
          confidence: values.length >= 2 ? 'medium' : 'manual',
          evidence: {
            source: values.length >= 2 ? 'repeated-examples' : 'document-structure',
            paragraphCount: values.length,
            note: values.length >= 2
              ? `Consistent across ${values.length} detected template example(s).`
              : 'Detected from one template example; manual confirmation is recommended.'
          }
        });
      } else {
        warnings.push(`${displayComponent(ruleComponent(rule))} ${rule.property} has inconsistent template examples.`);
      }
    }

    const measurableProperties: Array<[keyof Formatting, string]> = [
      ['fontName', 'font.name'],
      ['fontSize', 'font.size'],
      ['bold', 'font.bold'],
      ['italic', 'font.italic'],
      ['alignment', 'paragraph.alignment'],
      ['spaceBefore', 'paragraph.spaceBefore'],
      ['spaceAfter', 'paragraph.spaceAfter'],
      ['firstLineIndent', 'paragraph.firstLineIndent'],
      ['leftIndent', 'paragraph.leftIndent'],
      ['rightIndent', 'paragraph.rightIndent']
    ];
    for (const [component, style] of componentStyles) {
      for (const [formattingProperty, property] of measurableProperties) {
        if (candidateByKey.has(candidateKey(component, property))) continue;
        const detectedValue = style[formattingProperty];
        if (detectedValue === undefined) continue;
        addCandidate({
          component,
          property,
          detectedValue,
          confidence: 'high',
          evidence: {
            source: 'style',
            styleName: style.name,
            note: 'Measured from a named DOCX style with no matching current strategy rule.'
          }
        });
      }
    }

    for (const [key, candidate] of candidateByKey) {
      if (candidates.every(existing => candidateKey(existing.component, existing.property) !== key)) candidates.push(candidate);
    }

    const proposals: StrategyChangeProposal[] = [];
    for (const rule of strategy.rules) {
      const candidate = candidateByKey.get(candidateKey(ruleComponent(rule), rule.property));
      if (!candidate) {
        proposals.push({
          ruleId: rule.id,
          ruleName: rule.name,
          component: ruleComponent(rule),
          property: rule.property,
          currentValue: currentValueFor(rule),
          status: 'not-determinable',
          confidence: 'manual',
          note: 'This publication requirement cannot be established reliably from the uploaded template.'
        });
        continue;
      }
      const currentValue = currentValueFor(rule);
      const detectedSpacing = candidate.detectedValue && typeof candidate.detectedValue === 'object'
        && 'before' in candidate.detectedValue && 'after' in candidate.detectedValue
        ? candidate.detectedValue as { before: unknown; after: unknown }
        : undefined;
      const asymmetricSpacing = detectedSpacing !== undefined
        && sharedSpacingValue(detectedSpacing) === undefined;
      proposals.push({
        ruleId: rule.id,
        ruleName: rule.name,
        component: ruleComponent(rule),
        property: rule.property,
        currentValue,
        detectedValue: candidate.detectedValue,
        status: valuesEqual(currentValue, candidate.detectedValue) ? 'unchanged' : 'changed',
        confidence: asymmetricSpacing ? 'manual' : candidate.confidence,
        evidence: candidate.evidence,
        note: asymmetricSpacing
          ? 'Before/after spacing differs. The current strategy schema stores one shared spacing value, so this requires manual review.'
          : candidate.confidence === 'high'
          ? 'Directly measurable from the template style or document structure.'
          : 'Inferred from repeated examples; review before accepting.'
      });
    }

    const detectedKeys = new Set(candidates.map(candidate => candidateKey(candidate.component, candidate.property)));
    for (const candidate of candidates.filter(item => !item.ruleId)) {
      proposals.push({
        ruleId: `NEW_${candidate.component.toUpperCase()}_${candidate.property.replace(/[^A-Z0-9]+/gi, '_').toUpperCase()}`,
        ruleName: `New ${displayComponent(candidate.component)} ${candidate.property}`,
        component: candidate.component,
        property: candidate.property,
        detectedValue: candidate.detectedValue,
        status: 'new',
        confidence: candidate.confidence,
        evidence: candidate.evidence,
        note: 'No matching current strategy rule was found; create or configure this rule manually.'
      });
    }
    void detectedKeys;

    const counts = {
      unchanged: proposals.filter(proposal => proposal.status === 'unchanged').length,
      changed: proposals.filter(proposal => proposal.status === 'changed').length,
      new: proposals.filter(proposal => proposal.status === 'new').length,
      review: proposals.filter(proposal => proposal.confidence !== 'high' && proposal.status !== 'not-determinable').length,
      notDeterminable: proposals.filter(proposal => proposal.status === 'not-determinable').length
    };

    const components: TemplateComponentSummary[] = [];
    const knownComponents = new Set<string>([
      ...componentStyles.keys(),
      ...componentExamples.keys()
    ]);
    for (const component of knownComponents) {
      const style = componentStyles.get(component);
      const examples = componentExamples.get(component)?.length || 0;
      components.push({
        component,
        styleName: style?.name,
        examples,
        confidence: style ? 'high' : examples >= 2 ? 'medium' : 'manual'
      });
    }

    return { fileName, components, candidates, proposals, counts, warnings };
  }

  static applyAcceptedChanges(
    strategy: StrategyDefinition,
    analysis: TemplateAnalysisResult,
    acceptedRuleIds: Set<string>
  ): StrategyDefinition {
    const next: StrategyDefinition = JSON.parse(JSON.stringify(strategy)) as StrategyDefinition;
    for (const proposal of analysis.proposals) {
      if (!acceptedRuleIds.has(proposal.ruleId) || !isTemplateProposalAcceptable(proposal)) continue;
      const rule = next.rules.find(item => item.id === proposal.ruleId);
      if (!rule) continue;
      const detected = proposal.detectedValue;
      const sharedSpacing = sharedSpacingValue(detected);
      if (sharedSpacing !== undefined) {
        rule.expected = sharedSpacing;
      } else if (typeof detected === 'string' || typeof detected === 'number' || typeof detected === 'boolean') {
        rule.expected = detected;
      }
      rule.evidence = reviewedTemplateEvidence(proposal, analysis.fileName, detected, rule.unit);
    }
    next.version = `${strategy.version}-template-review`;
    next.source = `${strategy.source} Reviewed template: ${analysis.fileName}.`;
    return next;
  }
}
