import JSZip from 'jszip';
import { BlankParagraphInfo, DocumentBlockSnapshot, DocumentLayoutSnapshot, EffectiveTableBorders, FieldTextSpan, ParagraphLayoutSignals, ParagraphOoxmlSnapshot, ReferenceNumberingSnapshot, ReferenceStyleSnapshot, SectionLayoutSnapshot, TabStopSnapshot } from '../core/models';
import { paragraphTextFingerprint } from '../core/utils';
import {
  EffectiveRunFormatting,
  FormattingLayer,
  FormattingSource,
  LineSpacingValue,
  NumberingLevelFormatting,
  StyleFormatting,
  resolveEffectiveNumberingFormatting,
  resolveEffectiveFormatting,
  resolveEffectiveRunFormatting
} from './EffectiveFormattingResolver';
import { hasPageNumberField, isTableSequenceFieldInstruction } from './WordFieldAnalyzer';

export { hasPageNumberField } from './WordFieldAnalyzer';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function attr(el: Element | null, localName: string): string | null {
  if (!el) return null;
  return el.getAttributeNS(W_NS, localName) || el.getAttribute(`w:${localName}`) || el.getAttribute(localName);
}

function twipsToInches(value: string | null): number | undefined {
  if (!value) return undefined;
  const n = Number(value);
  return Number.isFinite(n) ? n / 1440 : undefined;
}

/**
 * XML 1.0 does not allow most C0 control characters. Word-generated OOXML
 * should not contain them, but some browser/file transfer paths can surface
 * stray control bytes after decoding. Removing only XML-forbidden controls
 * makes the parser more tolerant without changing document text/formatting.
 */
function sanitizeXml(xml: string): string {
  return xml
    .replace(/^\uFEFF/, '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}

function parserErrorText(doc: XMLDocument): string | null {
  const direct = doc.getElementsByTagName('parsererror')[0];
  if (direct) return direct.textContent?.trim() || 'XML parser error';

  // Firefox/other engines may namespace parsererror.
  const all = Array.from(doc.getElementsByTagName('*'));
  const namespaced = all.find(el => el.localName === 'parsererror');
  return namespaced?.textContent?.trim() || null;
}

function parseXmlWithRecovery(xml: string): XMLDocument {
  const parser = new DOMParser();
  const first = parser.parseFromString(xml, 'application/xml');
  if (!parserErrorText(first)) return first;

  const cleaned = sanitizeXml(xml);
  const second = parser.parseFromString(cleaned, 'application/xml');
  const error = parserErrorText(second);
  if (!error) return second;

  throw new Error(`Unable to parse OOXML part with DOMParser: ${error.slice(0, 180)}`);
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Namespace-prefix-agnostic attribute reader for a single OOXML start tag.
 * Example: attrFromTag('<w:pgMar w:top="1440"/>', 'top') => '1440'.
 */
function attrFromTag(tag: string | null, localName: string): string | null {
  if (!tag) return null;
  const name = escapeRegExp(localName);
  const pattern = new RegExp(`(?:^|\\s)(?:[A-Za-z_][\\w.-]*:)?${name}\\s*=\\s*["']([^"']+)["']`, 'i');
  return tag.match(pattern)?.[1] ?? null;
}

function elementStartTag(xml: string, localName: string): string | null {
  const name = escapeRegExp(localName);
  const match = xml.match(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>`, 'i'));
  return match?.[0] ?? null;
}

function allElementStartTags(xml: string, localName: string): string[] {
  const name = escapeRegExp(localName);
  return Array.from(xml.matchAll(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>`, 'gi')), m => m[0]);
}

function allElementBlocks(xml: string, localName: string): string[] {
  const name = escapeRegExp(localName);
  const blockPattern = new RegExp(
    `<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?${name}\\s*>`,
    'gi'
  );
  const blocks = Array.from(xml.matchAll(blockPattern), m => m[0]);

  // sectPr can technically be empty/self-closing, so include that shape too.
  const selfClosingPattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*/\\s*>`, 'gi');
  blocks.push(...Array.from(xml.matchAll(selfClosingPattern), m => m[0]));
  return blocks;
}

/** Return only direct child element blocks of the supplied XML root. */
function directChildBlocks(xml: string, localName: string): string[] {
  const wanted = localName.toLowerCase();
  const tokenPattern = /<\/?(?:[A-Za-z_][\w.-]*:)?[A-Za-z_][\w.-]*(?:\s[^<>]*?)?\/?\s*>/g;
  const blocks: string[] = [];
  let depth = 0;
  let captureStart: number | null = null;

  for (const match of xml.matchAll(tokenPattern)) {
    const token = match[0];
    const closing = /^<\//.test(token);
    const selfClosing = /\/\s*>$/.test(token);
    const local = token
      .replace(/^<\/?/, '')
      .match(/^(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)/i)?.[1]
      .toLowerCase();
    const start = match.index ?? 0;

    if (closing) {
      if (depth > 0) depth -= 1;
      if (captureStart !== null && local === wanted && depth === 1) {
        blocks.push(xml.slice(captureStart, start + token.length));
        captureStart = null;
      }
      continue;
    }

    if (depth === 1 && local === wanted) {
      if (selfClosing) blocks.push(token);
      else captureStart = start;
    }
    if (!selfClosing) depth += 1;
  }
  return blocks;
}

function countLocalElements(xml: string, localName: string): number {
  return allElementStartTags(xml, localName).length;
}

interface MainFlowParagraphBlock {
  xml: string;
  start: number;
  end: number;
}

/**
 * Return addressable Word paragraphs in document order. Paragraphs inside a
 * drawing/text-box are not part of body.paragraphs and must not shift the
 * indices used by Office.js Locate or paragraph formatting validation.
 */
function mainFlowParagraphBlocksWithOffsets(documentXml: string): MainFlowParagraphBlock[] {
  const tokenPattern = /<\/?(?:[A-Za-z_][\w.-]*:)?(p|txbxContent)\b[^>]*>/gi;
  const paragraphs: MainFlowParagraphBlock[] = [];
  let textBoxDepth = 0;
  let paragraphDepth = 0;
  let paragraphStart: number | null = null;

  for (const match of documentXml.matchAll(tokenPattern)) {
    const token = match[0];
    const localName = (match[1] || '').toLowerCase();
    const closing = /^<\//.test(token);
    const selfClosing = /\/\s*>$/.test(token);
    const start = match.index ?? 0;

    if (localName === 'txbxcontent') {
      if (closing) textBoxDepth = Math.max(0, textBoxDepth - 1);
      else if (!selfClosing) textBoxDepth += 1;
      continue;
    }

    if (textBoxDepth > 0) continue;

    if (closing) {
      if (paragraphDepth > 0) paragraphDepth -= 1;
      if (paragraphDepth === 0 && paragraphStart !== null) {
        paragraphs.push({
          xml: documentXml.slice(paragraphStart, start + token.length),
          start: paragraphStart,
          end: start + token.length
        });
        paragraphStart = null;
      }
      continue;
    }

    if (paragraphDepth === 0) paragraphStart = start;
    if (selfClosing) {
      paragraphs.push({ xml: token, start, end: start + token.length });
      paragraphStart = null;
    } else {
      paragraphDepth += 1;
    }
  }

  return paragraphs;
}

export interface ParagraphIndentRepairResult {
  /** Standalone target paragraph XML suitable for Word.Paragraph.insertOoxml. */
  xml: string;
  changed: boolean;
}

function removeIndentFromParagraphXml(paragraphXml: string): ParagraphIndentRepairResult {
  const pProperties = directChildBlocks(paragraphXml, 'pPr')[0];
  if (!pProperties) return { xml: paragraphXml, changed: false };
  const indentation = directChildBlocks(pProperties, 'ind')[0];
  if (!indentation) return { xml: paragraphXml, changed: false };
  const repairedProperties = pProperties.replace(indentation, '');
  const repairedParagraph = paragraphXml.replace(pProperties, repairedProperties);
  return { xml: repairedParagraph, changed: true };
}

function paragraphXmlText(paragraphXml: string): string {
  return paragraphTextFingerprint(
    allElementBlocks(paragraphXml, 'r').map(runText).join('')
  );
}

function domParagraphText(paragraph: Element): string {
  return paragraphTextFingerprint(
    Array.from(paragraph.getElementsByTagName('*'))
      .filter(element => element.localName === 't')
      .map(element => element.textContent || '')
      .join('')
  );
}

function hasRelationshipBackedContent(paragraphXml: string): boolean {
  // Relationship-bearing OOXML attributes can occur on hyperlinks, images,
  // OLE objects, AlternateContent, and extension elements. Do not restrict
  // this check to a known element list. A prefix may be declared on an
  // ancestor that is not present in an isolated paragraph payload, so an
  // unresolved prefixed id/embed/link is conservatively treated as unsafe.
  return /\b[A-Za-z_][\w.-]*:(?:id|embed|link)\s*=\s*["']/i.test(paragraphXml);
}

function chooseDomParagraph(
  paragraphs: Element[],
  paragraphIndex: number,
  targetText?: string
): Element | null {
  if (paragraphs.length === 1) return paragraphs[0];
  if (targetText !== undefined) {
    const expected = paragraphTextFingerprint(targetText);
    const matches = paragraphs.filter(paragraph => domParagraphText(paragraph) === expected);
    if (matches.length === 1) return matches[0];
  }
  // A document-wide index is not safe for an individual Flat OPC payload.
  // Refuse ambiguous packages instead of editing a different paragraph.
  return null;
}

function chooseXmlParagraph(
  paragraphs: MainFlowParagraphBlock[],
  paragraphIndex: number,
  targetText?: string
): MainFlowParagraphBlock | null {
  if (paragraphs.length === 1) return paragraphs[0];
  if (targetText !== undefined) {
    const expected = paragraphTextFingerprint(targetText);
    const matches = paragraphs.filter(paragraph => paragraphXmlText(paragraph.xml) === expected);
    if (matches.length === 1) return matches[0];
  }
  // A document-wide index is not safe for an individual Flat OPC payload.
  // Only a unique text match can identify one paragraph among several.
  return null;
}

function localAttribute(element: Element, localName: string): string | null {
  return element.getAttribute(localName)
    || element.getAttribute(`pkg:${localName}`)
    || element.getAttributeNS('http://schemas.microsoft.com/office/2006/xmlPackage', localName)
    || null;
}

function directChildElement(element: Element, localName: string): Element | null {
  return Array.from(element.children).find(child => child.localName === localName) || null;
}

function mainFlowDomParagraphs(root: Element): Element[] {
  return Array.from(root.getElementsByTagName('*'))
    .filter(element => element.localName === 'p')
    .filter(element => {
      let parent: Element | null = element.parentElement;
      while (parent) {
        if (parent.localName === 'txbxContent') return false;
        parent = parent.parentElement;
      }
      return true;
    });
}

function serializeElement(element: Element): string {
  if (typeof XMLSerializer !== 'undefined') return new XMLSerializer().serializeToString(element);
  return element.outerHTML;
}

function repairWithDomParser(ooxml: string, paragraphIndex: number, targetText?: string): ParagraphIndentRepairResult | null {
  if (typeof DOMParser === 'undefined') return null;
  const document = parseXmlWithRecovery(ooxml);
  const root = document.documentElement;
  let target: Element | null = null;

  if (root.localName === 'p') {
    target = root;
  } else if (root.localName === 'package') {
    const documentPart = Array.from(root.getElementsByTagName('*'))
      .find(element => element.localName === 'part'
        && localAttribute(element, 'name')?.toLowerCase() === '/word/document.xml');
    const xmlData = documentPart ? directChildElement(documentPart, 'xmlData') : null;
    const packageDocumentRoot = xmlData?.children[0] || null;
    if (packageDocumentRoot) {
      target = chooseDomParagraph(mainFlowDomParagraphs(packageDocumentRoot), paragraphIndex, targetText);
    }
  } else {
    target = chooseDomParagraph(mainFlowDomParagraphs(root), paragraphIndex, targetText);
  }

  if (!target) {
    throw new Error('The reference paragraph could not be located in the supplied OOXML.');
  }
  const pProperties = directChildElement(target, 'pPr');
  const indentation = pProperties ? directChildElement(pProperties, 'ind') : null;
  if (!indentation) {
    return { xml: ooxml, changed: false };
  }
  if (hasRelationshipBackedContent(serializeElement(target))) {
    throw new Error('Reference paragraph contains relationship-backed content such as a hyperlink. Indentation replacement was withheld for manual review.');
  }
  pProperties!.removeChild(indentation);
  // Flat OPC is a container, but insertOoxml on a Paragraph accepts only the
  // target w:p. Other package parts are left untouched by this operation.
  return { xml: serializeElement(target), changed: true };
}

function flatOpcDocumentXml(ooxml: string): string | null {
  const part = directChildBlocks(ooxml, 'part')
    .find(block => localAttributeFromTag(block, 'name')?.toLowerCase() === '/word/document.xml');
  if (!part) return null;
  const xmlData = directChildBlocks(part, 'xmlData')[0];
  if (!xmlData) return null;
  return xmlData
    .replace(/^<(?:[A-Za-z_][\w.-]*:)?xmlData\b[^>]*>/i, '')
    .replace(/<\/(?:[A-Za-z_][\w.-]*:)?xmlData\s*>\s*$/i, '');
}

function localAttributeFromTag(tag: string, localName: string): string | null {
  const escaped = escapeRegExp(localName);
  return tag.match(new RegExp(`(?:^|\\s)(?:[A-Za-z_][\\w.-]*:)?${escaped}\\s*=\\s*["']([^"']+)["']`, 'i'))?.[1] ?? null;
}

/**
 * Remove only the target paragraph's direct w:ind property. Office can return
 * a normal w:p or a Flat OPC package, so package styles must never be treated
 * as the target paragraph. The result is always standalone w:p XML for the
 * subsequent Word paragraph replacement.
 */
export function repairParagraphIndentOoxml(
  ooxml: string,
  paragraphIndex = 0,
  targetText?: string
): ParagraphIndentRepairResult {
  const domResult = repairWithDomParser(ooxml, paragraphIndex, targetText);
  if (domResult) return domResult;

  const packageDocument = /<(?:[A-Za-z_][\w.-]*:)?package\b/i.test(ooxml)
    ? flatOpcDocumentXml(ooxml)
    : null;
  const targetDocument = packageDocument || ooxml;
  const paragraphBlocks = mainFlowParagraphBlocksWithOffsets(targetDocument);
  const target = chooseXmlParagraph(paragraphBlocks, paragraphIndex, targetText);
  if (!target) throw new Error('The reference paragraph could not be located in the supplied OOXML.');
  const repaired = removeIndentFromParagraphXml(target.xml);
  if (repaired.changed && hasRelationshipBackedContent(target.xml)) {
    throw new Error('Reference paragraph contains relationship-backed content such as a hyperlink. Indentation replacement was withheld for manual review.');
  }
  return repaired.changed
    ? { xml: repaired.xml, changed: true }
    : { xml: ooxml, changed: false };
}

function stripTextBoxContent(xml: string): string {
  const tokenPattern = /<\/?(?:[A-Za-z_][\w.-]*:)?txbxContent\b[^>]*>/gi;
  const ranges: Array<{ start: number; end: number }> = [];
  let depth = 0;
  let rangeStart: number | null = null;

  for (const match of xml.matchAll(tokenPattern)) {
    const token = match[0];
    const closing = /^<\//.test(token);
    const selfClosing = /\/\s*>$/.test(token);
    const start = match.index ?? 0;
    if (closing) {
      if (depth > 0) depth -= 1;
      if (depth === 0 && rangeStart !== null) {
        ranges.push({ start: rangeStart, end: start + token.length });
        rangeStart = null;
      }
    } else if (!selfClosing) {
      if (depth === 0) rangeStart = start;
      depth += 1;
    } else {
      ranges.push({ start, end: start + token.length });
    }
  }

  if (!ranges.length) return xml;
  let result = '';
  let cursor = 0;
  for (const range of ranges) {
    result += xml.slice(cursor, range.start);
    cursor = range.end;
  }
  return result + xml.slice(cursor);
}

function decodeBasicXmlEntities(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)));
}

/**
 * An empty Word header/footer can still have a relationship and an XML part.
 * AIP forbids visible/meaningful header/footer content, not the existence of an
 * empty container. Inspect the referenced part instead of failing merely
 * because <w:headerReference>/<w:footerReference> exists.
 */
function hasMeaningfulHeaderFooterContent(xml: string): boolean {
  const clean = sanitizeXml(xml);

  // Text runs with non-whitespace content.
  const textRuns = Array.from(clean.matchAll(
    /<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t\s*>/gi
  ));
  if (textRuns.some(m => decodeBasicXmlEntities(m[1]).trim().length > 0)) return true;

  // Fields, drawings, VML, embedded objects, and content controls with media
  // are meaningful even when they contain no normal <w:t> text.
  if (/<(?:[A-Za-z_][\w.-]*:)?(?:instrText|fldSimple|drawing|pict|object)\b/i.test(clean)) return true;

  return false;
}

function normalizeWordPartTarget(target: string): string {
  const clean = target.replace(/\\/g, '/').replace(/^\//, '');
  const base = clean.toLowerCase().startsWith('word/') ? clean : `word/${clean}`;
  const components = base.split('/');
  const normalized: string[] = [];
  for (const component of components) {
    if (!component || component === '.') continue;
    if (component === '..') normalized.pop();
    else normalized.push(component);
  }
  return normalized.join('/');
}

function referencedHeaderFooterParts(
  documentXml: string,
  relationshipsXml: string,
  kind: 'header' | 'footer'
): string[] {
  const referenceIds = new Set(
    allElementStartTags(documentXml, `${kind}Reference`)
      .map(tag => attrFromTag(tag, 'id'))
      .filter((id): id is string => Boolean(id))
  );
  if (!referenceIds.size) return [];

  const parts: string[] = [];
  for (const tag of allElementStartTags(relationshipsXml, 'Relationship')) {
    const id = attrFromTag(tag, 'Id') || attrFromTag(tag, 'id');
    const target = attrFromTag(tag, 'Target') || attrFromTag(tag, 'target');
    const type = attrFromTag(tag, 'Type') || attrFromTag(tag, 'type') || '';
    if (!id || !target || !referenceIds.has(id)) continue;
    if (!new RegExp(`/${kind}$`, 'i').test(type) && !new RegExp(`${kind}\\d*\\.xml$`, 'i').test(target)) continue;
    parts.push(normalizeWordPartTarget(target));
  }
  return [...new Set(parts)];
}

interface ParsedLayout {
  sections: SectionLayoutSnapshot[];
  pageWidthInches?: number;
  pageHeightInches?: number;
  marginTopInches?: number;
  marginRightInches?: number;
  marginBottomInches?: number;
  marginLeftInches?: number;
  columns?: number;
  hasHeaderReference: boolean;
  hasFooterReference: boolean;
}

function removeMutableSectionProperties(xml: string): string {
  return xml
    .replace(/<(?:[A-Za-z_][\w.-]*:)?pgSz\b[^>]*\/\s*>/gi, '')
    .replace(/<(?:[A-Za-z_][\w.-]*:)?pgSz\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?pgSz\s*>/gi, '')
    .replace(/<(?:[A-Za-z_][\w.-]*:)?pgMar\b[^>]*\/\s*>/gi, '')
    .replace(/<(?:[A-Za-z_][\w.-]*:)?pgMar\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?pgMar\s*>/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function sectionBoundarySignatureFromDom(section: Element): string {
  const serializer = typeof XMLSerializer === 'undefined' ? null : new XMLSerializer();
  return Array.from(section.children)
    .filter(child => child.localName !== 'pgSz' && child.localName !== 'pgMar')
    .map(child => serializer?.serializeToString(child) || child.outerHTML || '')
    .join('')
    .replace(/\s+/g, ' ')
    .trim();
}

interface OoxmlStyleDefinition extends StyleFormatting {
  styleId?: string;
  name?: string;
  nextParagraphStyle?: string;
  numbering?: ParagraphNumberingProperties;
}

interface ParagraphNumberingProperties {
  numId: number;
  level: number;
}

interface ThemeFonts {
  majorLatin?: string;
  minorLatin?: string;
}

function parseThemeFonts(themeXml: string): ThemeFonts {
  const majorFont = allElementBlocks(themeXml, 'majorFont')[0] || '';
  const minorFont = allElementBlocks(themeXml, 'minorFont')[0] || '';
  const latinTypeface = (fontBlock: string): string | undefined =>
    attrFromTag(elementStartTag(fontBlock, 'latin'), 'typeface') || undefined;
  return {
    majorLatin: latinTypeface(majorFont),
    minorLatin: latinTypeface(minorFont)
  };
}

function themeFontName(fonts: string, themes: ThemeFonts): string | undefined {
  const theme = attrFromTag(fonts, 'asciiTheme')
    || attrFromTag(fonts, 'hAnsiTheme')
    || attrFromTag(fonts, 'eastAsiaTheme')
    || attrFromTag(fonts, 'csTheme');
  if (!theme) return undefined;
  if (/^major/i.test(theme)) return themes.majorLatin;
  if (/^minor/i.test(theme)) return themes.minorLatin;
  return undefined;
}

function parseRunFormatting(rPr: string, themes: ThemeFonts = {}): Pick<FormattingLayer, 'fontName' | 'fontSize' | 'bold' | 'italic' | 'allCaps' | 'verticalAlign' | 'baselinePositionPt'> {
  const fonts = elementStartTag(rPr, 'rFonts');
  const fontName = fonts
    ? themeFontName(fonts, themes)
      || attrFromTag(fonts, 'ascii')
      || attrFromTag(fonts, 'hAnsi')
      || attrFromTag(fonts, 'cs')
      || undefined
    : undefined;
  const halfPoints = numericAttribute(elementStartTag(rPr, 'sz'), 'val');
  const positionHalfPoints = numericAttribute(elementStartTag(rPr, 'position'), 'val');
  const verticalAlignValue = attrFromTag(elementStartTag(rPr, 'vertAlign'), 'val');
  return {
    fontName,
    fontSize: halfPoints === undefined ? undefined : halfPoints / 2,
    bold: booleanElement(rPr, 'b'),
    italic: booleanElement(rPr, 'i'),
    allCaps: booleanElement(rPr, 'caps'),
    baselinePositionPt: positionHalfPoints === undefined ? undefined : positionHalfPoints / 2,
    verticalAlign: verticalAlignValue === 'superscript' || verticalAlignValue === 'subscript'
      ? verticalAlignValue
      : undefined
  };
}

function runStyleId(rPr: string): string | null {
  return attrFromTag(elementStartTag(rPr, 'rStyle'), 'val');
}

function runText(runXml: string): string {
  // Preserve manual line breaks in the effective-run text. Without them, a
  // Word paragraph containing an affiliation line followed by a) / b) email
  // lines is flattened into one string and run offsets no longer identify the
  // email association markers correctly.
  const tokenPattern = /<(?:[A-Za-z_][\w.-]*:)?(?:t|br|cr|tab)\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?t\s*>)/gi;
  let text = '';
  for (const match of runXml.matchAll(tokenPattern)) {
    const token = match[0];
    if (/<(?:[A-Za-z_][\w.-]*:)?(?:br|cr)\b/i.test(token)) {
      text += '\n';
      continue;
    }
    if (/<(?:[A-Za-z_][\w.-]*:)?tab\b/i.test(token)) {
      text += '\t';
      continue;
    }
    const content = token.replace(/^<[\s\S]*?>|<\/[\s\S]*?>$/g, '');
    text += decodeBasicXmlEntities(content);
  }
  return text;
}

/**
 * Count ordinary manual line breaks that occur before the first meaningful
 * visible text in a paragraph.  A break after text is part of the paragraph's
 * internal layout and is deliberately not treated as an affiliation/email
 * separator.  When a field or drawing occurs before the first visible text,
 * the result is unknown so callers can withhold destructive edits safely.
 */
export function leadingManualLineBreakCount(paragraphXml: string): number | null {
  const xml = stripTextBoxContent(paragraphXml);
  const tokenPattern = /<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?t\s*>|<(?:[A-Za-z_][\w.-]*:)?(?:br|cr|tab)\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?(?:br|cr|tab)\s*>)|<(?:[A-Za-z_][\w.-]*:)?fldSimple\b[^>]*>|<(?:[A-Za-z_][\w.-]*:)?(?:fldChar|instrText)\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?(?:fldChar|instrText)\s*>)|<(?:[A-Za-z_][\w.-]*:)?(?:drawing|pict|object|sdt|txbxContent|customXml)\b[^>]*(?:\/\s*>|>)/gi;
  let count = 0;

  for (const match of xml.matchAll(tokenPattern)) {
    const token = match[0];
    const startTag = token.match(/^<[^>]+>/)?.[0] || token;

    if (/<(?:[A-Za-z_][\w.-]*:)?t\b/i.test(token)) {
      const content = token.replace(/^<[^>]*>|<\/[\s\S]*?>$/g, '');
      if (decodeBasicXmlEntities(content).trim().length > 0) return count;
      continue;
    }

    if (/<(?:[A-Za-z_][\w.-]*:)?tab\b/i.test(token)) continue;

    if (/<(?:[A-Za-z_][\w.-]*:)?(?:br|cr)\b/i.test(token)) {
      const breakType = attrFromTag(startTag, 'type')?.toLowerCase();
      if (breakType !== undefined && breakType !== null && breakType !== '' && breakType !== 'textwrapping') {
        return null;
      }
      count += 1;
      continue;
    }

    // Fields, drawings, content controls, and objects can carry visible
    // content that cannot be aligned reliably with an Office.js paragraph.
    return null;
  }

  return count;
}

function visibleTextLength(xml: string): number {
  const tokenPattern = /<(?:[A-Za-z_][\w.-]*:)?(?:t|br|cr|tab)\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?t\s*>)/gi;
  let length = 0;
  for (const match of xml.matchAll(tokenPattern)) {
    const token = match[0];
    if (/<(?:[A-Za-z_][\w.-]*:)?(?:br|cr)\b/i.test(token)) length += 1;
    else if (/<(?:[A-Za-z_][\w.-]*:)?tab\b/i.test(token)) length += 1;
    else length += decodeBasicXmlEntities(token.replace(/^<[^>]*>|<\/[^>]*>$/g, '')).length;
  }
  return length;
}

/** Map visible paragraph-text offsets occupied by complex/simple field results. */
function parseFieldTextSpans(paragraphXml: string): FieldTextSpan[] {
  const spans: FieldTextSpan[] = [];
  for (const match of paragraphXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?fldSimple\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?fldSimple\s*>/gi)) {
    const start = visibleTextLength(paragraphXml.slice(0, match.index ?? 0));
    const end = start + visibleTextLength(match[0]);
    const openingTag = match[0].match(/^<(?:[A-Za-z_][\w.-]*:)?fldSimple\b[^>]*>/i)?.[0] || '';
    const instruction = attrFromTag(openingTag, 'instr') || '';
    if (end > start) {
      spans.push({
        start,
        end,
        ...(isTableSequenceFieldInstruction(instruction) ? { kind: 'tableSequence' as const } : {})
      });
    }
  }

  type ActiveField = { separated: boolean; textStart: number | null; instructionParts: string[] };
  const active: ActiveField[] = [];
  // Keep self-closing fldChar tokens separate from following runs. A single
  // alternation with a generic `>` branch can otherwise consume the next
  // instrText/t element and hide the field separator/result.
  const tokenPattern = /<(?:[A-Za-z_][\w.-]*:)?fldChar\b[^>]*\/\s*>|<(?:[A-Za-z_][\w.-]*:)?instrText\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?instrText\s*>|<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?t\s*>|<(?:[A-Za-z_][\w.-]*:)?(?:br|cr|tab)\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?(?:br|cr|tab)\s*>)/gi;
  let textOffset = 0;
  for (const match of paragraphXml.matchAll(tokenPattern)) {
    const token = match[0];
    if (/<(?:[A-Za-z_][\w.-]*:)?fldChar\b/i.test(token)) {
      const fieldType = attrFromTag(token.match(/^<[^>]+>/)?.[0] || token, 'fldCharType')?.toLowerCase();
      if (fieldType === 'begin') active.push({ separated: false, textStart: null, instructionParts: [] });
      else if (fieldType === 'separate' && active.length) active[active.length - 1].separated = true;
      else if (fieldType === 'end' && active.length) {
        const field = active.pop()!;
        if (field.textStart !== null && textOffset > field.textStart) {
          spans.push({
            start: field.textStart,
            end: textOffset,
            ...(isTableSequenceFieldInstruction(field.instructionParts.join(''))
              ? { kind: 'tableSequence' as const }
              : {})
          });
        }
      }
      continue;
    }
    if (/<(?:[A-Za-z_][\w.-]*:)?(?:instrText)\b/i.test(token)) {
      const field = active[active.length - 1];
      if (field) {
        const openEnd = token.indexOf('>');
        const closeStart = token.lastIndexOf('</');
        if (openEnd >= 0 && closeStart > openEnd) {
          field.instructionParts.push(decodeBasicXmlEntities(token.slice(openEnd + 1, closeStart)));
        }
      }
      continue;
    }

    const length = /<(?:[A-Za-z_][\w.-]*:)?(?:br|cr|tab)\b/i.test(token)
      ? 1
      : decodeBasicXmlEntities(token.replace(/^<[^>]*>|<\/[^>]*>$/g, '')).length;
    const field = active[active.length - 1];
    if (field?.separated && field.textStart === null) field.textStart = textOffset;
    textOffset += length;
  }

  return spans
    .filter(span => span.end > span.start)
    .sort((a, b) => a.start - b.start)
    .filter((span, index, all) => index === 0 || span.start !== all[index - 1].start || span.end !== all[index - 1].end);
}

/**
 * Classify an apparently empty paragraph without confusing visual whitespace
 * with structural Word content. A paragraph containing only a manual break,
 * field, bookmark, list metadata, drawing, equation, or explicit page/section break is
 * deliberately not considered safe for automatic deletion.
 */
function inspectBlankParagraph(paragraphXml: string, text: string): BlankParagraphInfo {
  const whitespaceOnly = text
    .replace(/\u00A0/g, ' ')
    .trim()
    .length === 0;
  const hasExplicitPageBreak = /<(?:[A-Za-z_][\w.-]*:)?br\b[^>]*\btype\s*=\s*["']page["'][^>]*>/i.test(paragraphXml);
  const hasLastRenderedPageBreak = /<(?:[A-Za-z_][\w.-]*:)?lastRenderedPageBreak\b/i.test(paragraphXml);
  const hasManualBreak = /<(?:[A-Za-z_][\w.-]*:)?(?:br|cr)\b/i.test(paragraphXml);
  const hasSectionBreak = /<(?:[A-Za-z_][\w.-]*:)?sectPr\b/i.test(paragraphXml);
  const hasDrawing = /<(?:[A-Za-z_][\w.-]*:)?(?:drawing|pict|object)\b/i.test(paragraphXml);
  const hasField = /<(?:[A-Za-z_][\w.-]*:)?(?:fldChar|fldSimple|instrText)\b/i.test(paragraphXml);
  const hasEquation = /<(?:[A-Za-z_][\w.-]*:)?(?:oMath|oMathPara)\b/i.test(paragraphXml);
  const hasNumbering = /<(?:[A-Za-z_][\w.-]*:)?numPr\b/i.test(paragraphXml);
  const hasBookmarkContent = /<(?:[A-Za-z_][\w.-]*:)?(?:bookmarkStart|bookmarkEnd|sdt|customXml|permStart|permEnd)\b/i.test(paragraphXml);

  return {
    whitespaceOnly,
    hasManualBreak,
    hasExplicitPageBreak,
    hasLastRenderedPageBreak,
    hasSectionBreak,
    hasDrawing,
    hasField,
    hasEquation,
    hasNumbering,
    hasBookmarkContent,
    safeToDelete: whitespaceOnly
      && !hasManualBreak
      && !hasExplicitPageBreak
      && !hasSectionBreak
      && !hasDrawing
      && !hasField
      && !hasEquation
      && !hasNumbering
      && !hasBookmarkContent
  };
}

/** Exported for structural blank-paragraph regression tests. */
export { inspectBlankParagraph };

function visualAllCapsForText(text: string, allCaps: boolean | undefined): boolean | undefined {
  const letters = text.replace(/[^A-Za-z]/g, '');
  if (!letters) return undefined;
  return allCaps === true || letters === letters.toUpperCase();
}

function twipsToPoints(value: string | null): number | undefined {
  const n = value === null ? NaN : Number(value);
  return Number.isFinite(n) ? n / 20 : undefined;
}

function paragraphFormattingFromXml(pPr: string): FormattingLayer {
  const spacing = elementStartTag(pPr, 'spacing');
  const indentation = elementStartTag(pPr, 'ind');
  const alignment = elementStartTag(pPr, 'jc');
  const lineSpacing = lineSpacingFromXml(pPr);
  const firstLine = twipsToPoints(attrFromTag(indentation, 'firstLine'));
  const hanging = twipsToPoints(attrFromTag(indentation, 'hanging'));

  return {
    alignment: attrFromTag(alignment, 'val') || undefined,
    pageBreakBefore: booleanElement(pPr, 'pageBreakBefore'),
    firstLineIndent: firstLine !== undefined ? firstLine : hanging !== undefined ? -hanging : undefined,
    leftIndent: twipsToPoints(attrFromTag(indentation, 'left')),
    rightIndent: twipsToPoints(attrFromTag(indentation, 'right')),
    spaceBefore: twipsToPoints(attrFromTag(spacing, 'before')),
    spaceAfter: twipsToPoints(attrFromTag(spacing, 'after')),
    lineSpacing,
    tabStops: parseTabStops(pPr)
  };
}

function parseTabStops(pPr: string): TabStopSnapshot[] | undefined {
  const tags = allElementStartTags(pPr, 'tab');
  if (!tags.length) return undefined;
  const alignment = (value: string | null): TabStopSnapshot['alignment'] => {
    switch ((value || '').toLowerCase()) {
      case 'center': return 'center';
      case 'right': return 'right';
      case 'decimal': return 'decimal';
      case 'bar': return 'bar';
      case 'list': return 'list';
      case 'left': return 'left';
      default: return 'unknown';
    }
  };
  return tags
    .map(tag => ({
      positionPt: twipsToPoints(attrFromTag(tag, 'pos')),
      alignment: alignment(attrFromTag(tag, 'val'))
    }))
    .filter((tab): tab is TabStopSnapshot => tab.positionPt !== undefined)
    .map(tab => ({ positionPt: tab.positionPt as number, alignment: tab.alignment }));
}

function visibleBorder(tag: string | null): boolean {
  const value = attrFromTag(tag, 'val');
  return Boolean(tag) && (value === null || !/^(?:nil|none|0)$/i.test(value));
}

type TableBorderSide = 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV';
type TableBorderValues = Partial<Record<TableBorderSide, boolean>>;

interface TableStyleDefinition {
  basedOn?: string;
  borders: TableBorderValues;
}

function borderValue(xml: string, side: TableBorderSide): boolean | undefined {
  const tag = elementStartTag(xml, side);
  return tag ? visibleBorder(tag) : undefined;
}

function tableBorderValues(xml: string): TableBorderValues {
  const values: TableBorderValues = {};
  for (const side of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] as TableBorderSide[]) {
    const value = borderValue(xml, side);
    if (value !== undefined) values[side] = value;
  }
  return values;
}

function parseTableStyleDefinitions(stylesXml: string): Map<string, TableStyleDefinition> {
  const styles = new Map<string, TableStyleDefinition>();
  for (const block of allElementBlocks(stylesXml, 'style')) {
    const styleTag = elementStartTag(block, 'style');
    if ((attrFromTag(styleTag, 'type') || '').toLowerCase() !== 'table') continue;
    const styleId = attrFromTag(styleTag, 'styleId');
    if (!styleId) continue;
    const directTblPr = directChildBlocks(block, 'tblPr')[0] || '';
    const directBorders = allElementBlocks(directTblPr, 'tblBorders')[0] || '';
    const wholeTableStyle = directChildBlocks(block, 'tblStylePr')
      .find(stylePr => /^wholeTable$/i.test(attrFromTag(elementStartTag(stylePr, 'tblStylePr'), 'type') || ''));
    const wholeTablePr = wholeTableStyle ? directChildBlocks(wholeTableStyle, 'tblPr')[0] || '' : '';
    const wholeTableBorders = allElementBlocks(wholeTablePr, 'tblBorders')[0] || '';
    const directValues = tableBorderValues(directBorders);
    const wholeTableValues = tableBorderValues(wholeTableBorders);
    styles.set(styleId, {
      basedOn: attrFromTag(elementStartTag(block, 'basedOn'), 'val') || undefined,
      // wholeTable is more specific than the style-level tblPr. Merge it
      // property-by-property, while ignoring firstRow/lastRow/etc. rules.
      borders: { ...directValues, ...wholeTableValues }
    });
  }
  return styles;
}

function resolveTableStyleBorders(
  styleId: string | null,
  styles: Map<string, TableStyleDefinition>,
  seen = new Set<string>()
): { values: TableBorderValues; sources: Partial<Record<TableBorderSide, string>> } {
  if (!styleId || seen.has(styleId)) return { values: {}, sources: {} };
  const style = styles.get(styleId);
  if (!style) return { values: {}, sources: {} };

  const nextSeen = new Set(seen);
  nextSeen.add(styleId);
  const inherited = resolveTableStyleBorders(style.basedOn || null, styles, nextSeen);
  const values: TableBorderValues = { ...inherited.values };
  const sources: Partial<Record<TableBorderSide, string>> = { ...inherited.sources };
  for (const side of Object.keys(style.borders) as TableBorderSide[]) {
    values[side] = style.borders[side];
    sources[side] = `style:${styleId}`;
  }
  return { values, sources };
}

function tableStyleId(xml: string): string | null {
  const tblPr = allElementBlocks(xml, 'tblPr')[0] || '';
  return attrFromTag(elementStartTag(tblPr, 'tblStyle'), 'val');
}

function resolveEffectiveTableBorders(xml: string, tableStyles: Map<string, TableStyleDefinition>): EffectiveTableBorders {
  const styleId = tableStyleId(xml);
  const styleBorders = resolveTableStyleBorders(styleId, tableStyles);
  const tblPr = allElementBlocks(xml, 'tblPr')[0] || '';
  const directBorders = allElementBlocks(tblPr, 'tblBorders')[0] || '';
  const direct = tableBorderValues(directBorders);
  const values: Record<TableBorderSide, boolean | null> = {
    top: null,
    bottom: null,
    left: null,
    right: null,
    insideH: null,
    insideV: null
  };
  const sources: Partial<Record<TableBorderSide, string>> = {};
  for (const side of Object.keys(values) as TableBorderSide[]) {
    if (direct[side] !== undefined) {
      values[side] = direct[side] as boolean;
      sources[side] = 'direct';
    } else if (styleBorders.values[side] !== undefined) {
      values[side] = styleBorders.values[side] as boolean;
      sources[side] = styleBorders.sources[side] || `style:${styleId}`;
    }
  }
  return { ...values, sources };
}

function rowBorderVisible(row: string, side: 'top' | 'bottom'): boolean {
  const trPr = allElementBlocks(row, 'trPr')[0] || '';
  const borders = allElementBlocks(trPr, 'tblBorders')[0] || '';
  return visibleBorder(elementStartTag(borders, side));
}

function allCellsHaveBorder(row: string, side: 'left' | 'right' | 'top' | 'bottom'): boolean {
  const cells = allElementBlocks(row, 'tc');
  if (!cells.length) return false;
  return cells.every(cell => {
    const borders = allElementBlocks(cell, 'tcBorders')[0] || '';
    return visibleBorder(elementStartTag(borders, side));
  });
}

function anyCellHasBorder(row: string, side: 'left' | 'right' | 'top' | 'bottom'): boolean {
  return allElementBlocks(row, 'tc').some(cell => {
    const borders = allElementBlocks(cell, 'tcBorders')[0] || '';
    return visibleBorder(elementStartTag(borders, side));
  });
}

function rowBoundaryFullyVisible(upper: string, lower: string): boolean {
  return rowBorderVisible(upper, 'bottom')
    || rowBorderVisible(lower, 'top')
    || allCellsHaveBorder(upper, 'bottom')
    || allCellsHaveBorder(lower, 'top');
}

function rowBoundaryAnyVisible(upper: string, lower: string): boolean {
  return rowBorderVisible(upper, 'bottom')
    || rowBorderVisible(lower, 'top')
    || anyCellHasBorder(upper, 'bottom')
    || anyCellHasBorder(lower, 'top');
}

function borderState(xml: string, name: 'top' | 'bottom', effectiveValue?: boolean | null): boolean | null {
  const tblPr = allElementBlocks(xml, 'tblPr')[0] || '';
  const borders = allElementBlocks(tblPr, 'tblBorders')[0] || '';
  const directValue = borderValue(borders, name);
  const rows = allElementBlocks(xml, 'tr');
  if (!rows.length) {
    if (directValue !== undefined) return directValue;
    if (effectiveValue !== undefined && effectiveValue !== null) return effectiveValue;
    return borders ? false : null;
  }
  const edgeRow = name === 'top' ? rows[0] : rows[rows.length - 1];
  if (rowBorderVisible(edgeRow, name)) return true;
  const cells = allElementBlocks(edgeRow, 'tc');
  if (cells.length && cells.every(cell => {
    const cellBorders = allElementBlocks(cell, 'tcBorders')[0] || '';
    return visibleBorder(elementStartTag(cellBorders, name));
  })) return true;
  if (directValue !== undefined) return directValue;
  if (effectiveValue !== undefined && effectiveValue !== null) return effectiveValue;
  return cells.length ? false : (borders ? false : null);
}

interface TableRuleState {
  rule: boolean | null;
  rows: number | null;
  vertical: boolean | null;
  extraHorizontal: boolean | null;
  captionRowIndex: number | null;
  headerStartRowIndex: number | null;
  headerEndRowIndex: number | null;
  topRule: boolean | null;
}

function rowCells(row: string): string[] {
  return allElementBlocks(row, 'tc');
}

function rowHasGridSpan(row: string): boolean {
  return rowCells(row).some(cell => {
    const tcPr = allElementBlocks(cell, 'tcPr')[0] || '';
    const tag = elementStartTag(tcPr, 'gridSpan');
    return tag !== undefined && Number(attrFromTag(tag, 'val') || '1') > 1;
  });
}

function rowHasHeaderLikeText(row: string): boolean {
  const texts = rowCells(row)
    .map(cell => blockText(cell).trim())
    .filter(Boolean);
  return texts.length >= 2 && texts.filter(text => /[A-Za-z]/.test(text)).length >= 2;
}

function hasVerticalMergeContinuation(upper: string, lower: string): boolean {
  const upperCells = rowCells(upper);
  const lowerCells = rowCells(lower);
  const count = Math.min(upperCells.length, lowerCells.length);
  for (let index = 0; index < count; index += 1) {
    const upperTcPr = allElementBlocks(upperCells[index], 'tcPr')[0] || '';
    const lowerTcPr = allElementBlocks(lowerCells[index], 'tcPr')[0] || '';
    const upperMerge = elementStartTag(upperTcPr, 'vMerge');
    const lowerMerge = elementStartTag(lowerTcPr, 'vMerge');
    if (!upperMerge || !lowerMerge) continue;
    const upperValue = attrFromTag(upperMerge, 'val');
    const lowerValue = attrFromTag(lowerMerge, 'val');
    if ((upperValue === null || upperValue === 'restart')
      && (lowerValue === null || lowerValue === 'continue')) return true;
  }
  return false;
}

function isLikelyHeaderContinuation(upper: string, lower: string): boolean {
  if (hasVerticalMergeContinuation(upper, lower)) return true;
  const upperCells = rowCells(upper);
  const lowerCells = rowCells(lower);
  return rowHasGridSpan(upper)
    && lowerCells.length > upperCells.length
    && rowHasHeaderLikeText(lower);
}

function inferHeaderEndRow(rows: string[], start: number, end: number): number {
  let inferredEnd = end;
  while (inferredEnd + 1 < rows.length - 1
    && isLikelyHeaderContinuation(rows[inferredEnd], rows[inferredEnd + 1])) {
    inferredEnd += 1;
  }
  return inferredEnd;
}

function tableHeaderRuleState(xml: string, effectiveBorders?: EffectiveTableBorders): TableRuleState {
  const rows = allElementBlocks(xml, 'tr');
  if (!rows.length) return { rule: null, rows: null, vertical: null, extraHorizontal: null, captionRowIndex: null, headerStartRowIndex: null, headerEndRowIndex: null, topRule: null };
  const tablePr = allElementBlocks(xml, 'tblPr')[0] || '';
  const tableBorders = allElementBlocks(tablePr, 'tblBorders')[0] || '';
  // An embedded AIP caption is only structurally valid when it is the first
  // logical row.  A later TABLE n. paragraph must not reclassify the table's
  // real header/body boundaries.
  const candidateCaptionRowIndex = rows.findIndex(row => /^TABLE\s+(?:\d+[A-Z]?|[IVXLCDM]+)\s*[.:-]/i.test(blockText(row).trim()));
  const captionRowIndex = candidateCaptionRowIndex === 0 ? 0 : -1;
  const isHeaderRow = (row: string): boolean => {
    const trPr = allElementBlocks(row, 'trPr')[0] || '';
    const tag = elementStartTag(trPr, 'tblHeader');
    if (!tag) return false;
    const value = attrFromTag(tag, 'val');
    return value === null || !/^(?:0|false|off)$/i.test(value);
  };
  const explicitHeaderIndices = rows
    .map((row, index) => isHeaderRow(row) ? index : -1)
    .filter(index => index >= 0 && index !== captionRowIndex);
  const fallbackHeaderIndex = captionRowIndex >= 0 ? captionRowIndex + 1 : 0;
  const headerIndices = explicitHeaderIndices.length
    ? explicitHeaderIndices
    : rows.length > fallbackHeaderIndex + 1 ? [fallbackHeaderIndex] : [];
  const headerStartRowIndex = headerIndices.length ? headerIndices[0] : null;
  const detectedHeaderEndRowIndex = headerIndices.length ? headerIndices[headerIndices.length - 1] : null;
  const headerEndRowIndex = detectedHeaderEndRowIndex === null
    ? null
    : inferHeaderEndRow(rows, headerStartRowIndex!, detectedHeaderEndRowIndex);
  const cellVertical = rows.some(row => anyCellHasBorder(row, 'left') || anyCellHasBorder(row, 'right'));
  const effectiveVertical = effectiveBorders
    ? [effectiveBorders.insideV, effectiveBorders.left, effectiveBorders.right]
      .some(value => value === true)
    : false;
  const knownVertical = effectiveBorders
    ? [effectiveBorders.insideV, effectiveBorders.left, effectiveBorders.right]
      .some(value => value !== null) || rows.length > 0
    : Boolean(tableBorders);
  const vertical = cellVertical || effectiveVertical
    ? true
    : knownVertical
      ? false
      : null;
  const tableInsideHorizontal = effectiveBorders
    ? effectiveBorders.insideH === true
    : visibleBorder(elementStartTag(tableBorders, 'insideH'));
  const firstBodyIndex = headerEndRowIndex === null ? -1 : headerEndRowIndex + 1;
  const bodyRowCount = firstBodyIndex >= 0 ? rows.length - firstBodyIndex : rows.length;
  let bodyHasHorizontal = false;
  let headerHasHorizontal = false;
  if (headerStartRowIndex !== null && headerEndRowIndex !== null) {
    for (let lowerIndex = headerStartRowIndex + 1; lowerIndex <= headerEndRowIndex; lowerIndex++) {
      if (rowBoundaryAnyVisible(rows[lowerIndex - 1], rows[lowerIndex])) {
        headerHasHorizontal = true;
        break;
      }
    }
  }
  if (firstBodyIndex >= 0) {
    for (let lowerIndex = firstBodyIndex + 1; lowerIndex < rows.length; lowerIndex++) {
      if (rowBoundaryAnyVisible(rows[lowerIndex - 1], rows[lowerIndex])) {
        bodyHasHorizontal = true;
        break;
      }
    }
  }
  const hasInternalHeaderBoundary = headerStartRowIndex !== null
    && headerEndRowIndex !== null
    && headerEndRowIndex > headerStartRowIndex;
  const insideHCreatesExtra = tableInsideHorizontal
    && (hasInternalHeaderBoundary || bodyRowCount > 1);
  const extraHorizontal = headerHasHorizontal
    || insideHCreatesExtra
    || bodyHasHorizontal;
  const rule = headerEndRowIndex !== null
    && firstBodyIndex < rows.length
    && (tableInsideHorizontal || rowBoundaryFullyVisible(rows[headerEndRowIndex], rows[firstBodyIndex]));
  const topRule = captionRowIndex >= 0 && headerStartRowIndex !== null
    ? (tableInsideHorizontal || rowBoundaryFullyVisible(rows[captionRowIndex], rows[headerStartRowIndex]))
    : borderState(xml, 'top', effectiveBorders?.top);
  return {
    rule: headerIndices.length ? rule : null,
    rows: headerStartRowIndex !== null && headerEndRowIndex !== null
      ? headerEndRowIndex - headerStartRowIndex + 1
      : 0,
    vertical,
    extraHorizontal,
    captionRowIndex: captionRowIndex >= 0 ? captionRowIndex : null,
    headerStartRowIndex,
    headerEndRowIndex,
    topRule
  };
}

function paragraphVisualAllCaps(runs: EffectiveRunFormatting[]): boolean | undefined {
  const results = runs
    .map(run => visualAllCapsForText(run.text, run.allCaps === null ? undefined : run.allCaps))
    .filter((value): value is boolean => value !== undefined);

  return results.length ? results.every(Boolean) : undefined;
}

function captionPrefixVisualAllCaps(
  runs: EffectiveRunFormatting[],
  kind: 'figure' | 'table'
): boolean | undefined {
  const paragraphText = runs.map(run => run.text).join('');
  const captionNumber = '(?:\\d+[A-Za-z]?|[IVXLCDM]+)';
  const pattern = kind === 'figure'
    ? new RegExp(`^\\s*(FIGURE\\s+${captionNumber}\\s*\\.)`, 'i')
    : new RegExp(`^\\s*(TABLE\\s+${captionNumber}\\s*\\.)`, 'i');
  const match = paragraphText.match(pattern);
  if (!match) return undefined;

  const prefixLength = match[0].length;
  let offset = 0;
  const results: boolean[] = [];
  for (const run of runs) {
    const text = run.text;
    const start = Math.max(0, Math.min(text.length, prefixLength - offset));
    const prefixPart = text.slice(0, start);
    if (prefixPart) {
      const visual = visualAllCapsForText(prefixPart, run.allCaps === null ? undefined : run.allCaps);
      if (visual !== undefined) results.push(visual);
    }
    offset += text.length;
    if (offset >= prefixLength) break;
  }

  return results.length ? results.every(Boolean) : undefined;
}

function numericAttribute(tag: string | null, name: string): number | undefined {
  const raw = attrFromTag(tag, name);
  if (raw === null || raw.trim() === '') return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function booleanElement(xml: string, name: string): boolean | undefined {
  const tag = elementStartTag(xml, name);
  if (!tag) return undefined;
  const value = attrFromTag(tag, 'val');
  if (value === null) return true;
  return !/^(?:0|false|off|none)$/i.test(value);
}

function lineSpacingFromXml(xml: string): LineSpacingValue | undefined {
  const tag = elementStartTag(xml, 'spacing');
  if (!tag) return undefined;

  const rawLine = attrFromTag(tag, 'line');
  const rawRule = attrFromTag(tag, 'lineRule');
  // w:spacing also stores paragraph before/after values. Those do not
  // override line spacing and must not shadow an inherited style value.
  if (rawLine === null && rawRule === null) return undefined;

  const line = numericAttribute(tag, 'line');
  const rule = (rawRule || 'auto').toLowerCase();
  return {
    multiple: rule === 'auto' ? (line === undefined ? 1 : line / 240) : null,
    rule,
    line: line ?? null
  };
}

function parseStyleDefinition(styleBlock: string, themes: ThemeFonts): OoxmlStyleDefinition {
  const styleTag = elementStartTag(styleBlock, 'style');
  const styleId = attrFromTag(styleTag, 'styleId') || undefined;
  const basedOn = attrFromTag(elementStartTag(styleBlock, 'basedOn'), 'val') || undefined;
  const name = attrFromTag(elementStartTag(styleBlock, 'name'), 'val') || undefined;
  const nextParagraphStyle = attrFromTag(elementStartTag(styleBlock, 'next'), 'val') || undefined;
  const pPr = allElementBlocks(styleBlock, 'pPr')[0] || '';
  const rPr = allElementBlocks(styleBlock, 'rPr')[0] || '';
  return {
    styleId,
    name,
    nextParagraphStyle,
    type: attrFromTag(styleTag, 'type') || undefined,
    basedOn,
    numbering: paragraphNumberingProperties(pPr),
    ...paragraphFormattingFromXml(pPr),
    ...parseRunFormatting(rPr, themes)
  };
}

interface ParsedNumberingDefinition {
  numId: number;
  levels: Map<number, NumberingLevelFormatting>;
}

function numberAttribute(xml: string, elementName: string, attributeName = 'val'): number | null {
  const value = attrFromTag(elementStartTag(xml, elementName), attributeName);
  if (value === null || value.trim() === '') return null;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseNumberingDefinitions(numberingXml: string): Map<number, ParsedNumberingDefinition> {
  const abstractLevels = new Map<number, Map<number, NumberingLevelFormatting>>();
  for (const abstractBlock of allElementBlocks(numberingXml, 'abstractNum')) {
    const abstractId = numberAttribute(abstractBlock, 'abstractNum', 'abstractNumId');
    if (abstractId === null) continue;
    const levels = new Map<number, NumberingLevelFormatting>();
    for (const levelBlock of allElementBlocks(abstractBlock, 'lvl')) {
      const level = numberAttribute(levelBlock, 'lvl', 'ilvl');
      if (level === null) continue;
      const pPr = allElementBlocks(levelBlock, 'pPr')[0] || '';
      const left = numberAttribute(pPr, 'ind', 'left');
      const hanging = numberAttribute(pPr, 'ind', 'hanging');
      // Number position is the list indentation anchor, not the optional tab
      // stop used after the generated label. A tab stop at 0.5 in is common
      // even when the number itself is aligned at 0.25 in.
      const numberPosition = left !== null && hanging !== null
        ? (left - hanging) / 20
        : null;
      const textPosition = left === null ? null : left / 20;
      const levelText = attrFromTag(elementStartTag(levelBlock, 'lvlText'), 'val');
      const suffix = levelText?.match(/([.。:;)])$/)?.[1] || null;
      const separator = attrFromTag(elementStartTag(levelBlock, 'suff'), 'val') || null;
      levels.set(level, resolveEffectiveNumberingFormatting({
        format: attrFromTag(elementStartTag(levelBlock, 'numFmt'), 'val'),
        suffix,
        separator,
        startAt: numberAttribute(levelBlock, 'start'),
        alignment: attrFromTag(elementStartTag(levelBlock, 'lvlJc'), 'val'),
        numberPositionPt: numberPosition,
        textPositionPt: textPosition
      }));
    }
    abstractLevels.set(abstractId, levels);
  }

  const definitions = new Map<number, ParsedNumberingDefinition>();
  for (const numBlock of allElementBlocks(numberingXml, 'num')) {
    const numId = numberAttribute(numBlock, 'num', 'numId');
    const abstractId = numberAttribute(numBlock, 'abstractNumId');
    if (numId === null || abstractId === null) continue;
    const levels = new Map<number, NumberingLevelFormatting>(abstractLevels.get(abstractId) || []);
    for (const override of allElementBlocks(numBlock, 'lvlOverride')) {
      const level = numberAttribute(override, 'lvlOverride', 'ilvl');
      if (level === null) continue;
      const startOverride = numberAttribute(override, 'startOverride');
      const current = levels.get(level);
      if (current && startOverride !== null) levels.set(level, { ...current, startAt: startOverride });
    }
    definitions.set(numId, { numId, levels });
  }
  return definitions;
}

function paragraphNumberingProperties(pPr: string): ParagraphNumberingProperties | undefined {
  const numPr = allElementBlocks(pPr, 'numPr')[0] || '';
  const numId = numberAttribute(numPr, 'numId');
  if (numId === null) return undefined;
  return {
    numId,
    level: numberAttribute(numPr, 'ilvl') ?? 0
  };
}

function referenceStyleSnapshot(
  styleId: string | null,
  styles: Map<string, OoxmlStyleDefinition>,
  defaults: FormattingLayer
): ReferenceStyleSnapshot | undefined {
  if (!styleId) return undefined;
  const style = styles.get(styleId);
  if (!style || !/(?:^|\b)reference(?:$|\b)/i.test(style.name || styleId)) return undefined;
  const effective = resolveEffectiveFormatting({}, styleId, styles, defaults);
  return {
    name: style.name || styleId,
    type: style.type || null,
    basedOn: style.basedOn || null,
    nextParagraphStyle: style.nextParagraphStyle || null,
    fontName: effective.fontName,
    fontSize: effective.fontSize,
    alignment: effective.alignment,
    // Word stores a hanging paragraph style as the total left position plus
    // a negative first-line indent.  The Reference style's user-facing left
    // indent is therefore the first-line anchor, not the raw w:left value.
    // In the AIP template both values are 426 twips, so the effective left
    // indent is 0 pt even though the raw XML contains 21.3 pt for w:left.
    leftIndent: effective.leftIndent === null
      ? null
      : effective.leftIndent - (effective.firstLineIndent === null
        ? 0
        : Math.max(0, -effective.firstLineIndent)),
    hangingIndent: effective.firstLineIndent === null ? null : Math.max(0, -effective.firstLineIndent)
  };
}

function paragraphReferenceNumbering(
  paragraphXml: string,
  numbering: Map<number, ParsedNumberingDefinition>,
  styleId: string | null,
  styles: Map<string, OoxmlStyleDefinition>
): ReferenceNumberingSnapshot {
  const pPr = allElementBlocks(paragraphXml, 'pPr')[0] || '';
  let properties = paragraphNumberingProperties(pPr);
  let source: ReferenceNumberingSnapshot['source'] = properties ? 'direct' : undefined;
  if (!properties) {
    let current = styleId;
    const seen = new Set<string>();
    while (current && !seen.has(current)) {
      seen.add(current);
      const style = styles.get(current);
      if (!style) break;
      if (style.numbering) {
        properties = style.numbering;
        source = 'style';
        break;
      }
      current = style.basedOn || null;
    }
  }
  if (!properties) {
    return {
      kind: 'none', level: null, format: null, suffix: null, startAt: null,
      separator: null, alignment: null, source: 'none', numberPositionPt: null, textPositionPt: null
    };
  }
  const definition = numbering.get(properties.numId);
  const formatting = definition?.levels.get(properties.level);
  return {
    kind: 'automatic',
    level: properties.level + 1,
    format: formatting?.format || null,
    suffix: formatting?.suffix || null,
    separator: formatting?.separator || null,
    alignment: formatting?.alignment || null,
    source,
    startAt: formatting?.startAt ?? null,
    numberPositionPt: formatting?.numberPositionPt ?? null,
    textPositionPt: formatting?.textPositionPt ?? null
  };
}

export { parseNumberingDefinitions };

interface ParagraphTableMetadata {
  containingTableIndex: number | null;
  containingTableRowIndex: number | null;
  tableNestingLevel: number;
}

function paragraphTableMetadata(documentXml: string): ParagraphTableMetadata[] {
  const metadata: ParagraphTableMetadata[] = [];
  const tokenPattern = /<\/?(?:[A-Za-z_][\w.-]*:)?(tbl|tr|p|txbxContent)\b[^>]*>/gi;
  const tables: Array<{ topLevelIndex: number; rowIndex: number }> = [];
  let nextTopLevelTableIndex = 0;
  let textBoxDepth = 0;
  let match: RegExpExecArray | null;
  while ((match = tokenPattern.exec(documentXml))) {
    const localName = (match[1] || match[2]).toLowerCase();
    const token = match[0];
    const closing = token.startsWith('</');
    const selfClosing = /\/\s*>$/.test(token);
    if (localName === 'txbxcontent') {
      if (closing) textBoxDepth = Math.max(0, textBoxDepth - 1);
      else if (!selfClosing) textBoxDepth += 1;
      continue;
    }
    if (textBoxDepth > 0) continue;
    if (closing) {
      if (localName === 'tbl') tables.pop();
      continue;
    }
    if (localName === 'tbl') {
      const topLevelIndex = tables.length ? tables[0].topLevelIndex : nextTopLevelTableIndex++;
      tables.push({ topLevelIndex, rowIndex: -1 });
    } else if (localName === 'tr') {
      const current = tables[tables.length - 1];
      if (current) current.rowIndex += 1;
    } else if (localName === 'p') {
      const current = tables[tables.length - 1];
      metadata.push({
        containingTableIndex: current?.topLevelIndex ?? null,
        containingTableRowIndex: current?.rowIndex ?? null,
        tableNestingLevel: tables.length
      });
    }
  }
  return metadata;
}

function isCaptionCandidateText(text: string): boolean {
  return /^\s*(?:TABLE|FIGURE|FIG\.?)\s+(?:\d+[A-Za-z]?|[IVXLCDM]+)\s*[.:-]/i.test(text);
}

/**
 * Manually constructed tables in Word are often a run of tabbed/drawn
 * paragraphs rather than a w:tbl. Protect their dense row region from
 * generic prose formatting and whitespace corrections.
 */
function markPseudoTableLayout(
  paragraphs: ParagraphOoxmlSnapshot[],
  tableMetadata: ParagraphTableMetadata[]
): ParagraphOoxmlSnapshot[] {
  const isLayoutSignal = (paragraph: ParagraphOoxmlSnapshot): boolean =>
    paragraph.layoutSignals?.hasDrawing === true
    || paragraph.layoutSignals?.hasAnchoredDrawing === true
    || (paragraph.layoutSignals?.hasCustomTabs === true && paragraph.layoutSignals?.hasNumbering !== true);

  const protectedIndices = new Set<number>();
  paragraphs.forEach(paragraph => {
    if (paragraph.layoutSignals?.hasDrawing || paragraph.layoutSignals?.hasAnchoredDrawing) {
      protectedIndices.add(paragraph.index);
    }
  });

  paragraphs.forEach((caption, captionIndex) => {
    if (!isCaptionCandidateText(caption.text || '') || (tableMetadata[captionIndex]?.tableNestingLevel || 0) > 0) return;

    const end = Math.min(paragraphs.length, captionIndex + 40);
    const signals: ParagraphOoxmlSnapshot[] = [];
    let ordinaryParagraphsAfterCluster = 0;
    for (const paragraph of paragraphs.slice(captionIndex + 1, end)) {
      if ((tableMetadata[paragraph.index]?.tableNestingLevel || 0) > 0) continue;
      if (isCaptionCandidateText(paragraph.text || '')) break;

      if (isLayoutSignal(paragraph)) {
        signals.push(paragraph);
        ordinaryParagraphsAfterCluster = 0;
        continue;
      }

      // Blank separators can occur inside a manually laid-out table. Two
      // consecutive meaningful non-layout paragraphs, however, are strong
      // evidence that ordinary manuscript prose has resumed.
      if ((paragraph.text || '').trim().length > 0) {
        ordinaryParagraphsAfterCluster += 1;
        if (ordinaryParagraphsAfterCluster >= 2) break;
      }
    }

    // A single tabbed paragraph is commonly a list or a reference. Require a
    // cluster before treating the following content as a pseudo-table.
    if (signals.length < 2) return;
    const firstSignal = signals[0].index;
    const lastSignal = signals[signals.length - 1].index;
    for (let index = firstSignal; index <= lastSignal; index += 1) {
      if ((tableMetadata[index]?.tableNestingLevel || 0) === 0) protectedIndices.add(index);
    }
  });

  return paragraphs.map(paragraph => ({
    ...paragraph,
    layoutSensitive: protectedIndices.has(paragraph.index) || paragraph.layoutSensitive === true
  }));
}

function parseParagraphOoxml(documentXml: string, stylesXml: string, themeXml = '', numberingXml = ''): ParagraphOoxmlSnapshot[] {
  const themes = parseThemeFonts(themeXml);
  const numbering = parseNumberingDefinitions(numberingXml);
  const tableMetadata = paragraphTableMetadata(documentXml);
  const paragraphBlocks = mainFlowParagraphBlocksWithOffsets(documentXml);
  const styles = new Map<string, OoxmlStyleDefinition>();
  let defaultParagraphStyleId: string | null = null;
  for (const block of allElementBlocks(stylesXml, 'style')) {
    const styleTag = elementStartTag(block, 'style');
    const styleId = attrFromTag(styleTag, 'styleId');
    const type = attrFromTag(styleTag, 'type');
    if (styleId) styles.set(styleId, parseStyleDefinition(block, themes));
    if (styleId && type === 'paragraph' && attrFromTag(styleTag, 'default') === '1') {
      defaultParagraphStyleId = styleId;
    }
  }
  const defaultsBlock = allElementBlocks(stylesXml, 'docDefaults')[0] || '';
  const defaultParagraphProperties = allElementBlocks(defaultsBlock, 'pPrDefault')[0] || '';
  const defaultPPr = allElementBlocks(defaultParagraphProperties, 'pPr')[0] || '';
  const defaults: FormattingLayer = {
    ...paragraphFormattingFromXml(defaultPPr),
    ...parseRunFormatting(allElementBlocks(defaultsBlock, 'rPrDefault')[0]
      ? allElementBlocks(allElementBlocks(defaultsBlock, 'rPrDefault')[0], 'rPr')[0] || ''
      : '', themes)
  };

  const parsed = paragraphBlocks.map((paragraphBlock, index) => {
    const paragraphXml = stripTextBoxContent(paragraphBlock.xml);
    const pPr = allElementBlocks(paragraphXml, 'pPr')[0] || '';
    const direct: FormattingLayer = {
      ...paragraphFormattingFromXml(pPr),
      ...parseRunFormatting(allElementBlocks(pPr, 'rPr')[0] || '', themes)
    };
    const explicitStyleId = attrFromTag(elementStartTag(pPr, 'pStyle'), 'val');
    const styleId = explicitStyleId || defaultParagraphStyleId;
    const effective = resolveEffectiveFormatting(direct, styleId, styles, defaults);
    const referenceStyle = referenceStyleSnapshot(styleId, styles, defaults);
    const referenceNumbering = paragraphReferenceNumbering(paragraphXml, numbering, styleId, styles);
    const source = effective.sources.lineSpacingMultiple as ParagraphOoxmlSnapshot['lineSpacingSource'];
    const effectiveRuns = allElementBlocks(paragraphXml, 'r')
      .map(runXml => {
        const text = runText(runXml);
        if (!text) return null;
        const rPr = allElementBlocks(runXml, 'rPr')[0] || '';
        return resolveEffectiveRunFormatting(
          text,
          parseRunFormatting(rPr, themes),
          runStyleId(rPr),
          effective,
          styles,
          defaults
        );
      })
      .filter((run): run is EffectiveRunFormatting => run !== null);
    const paragraphText = allElementBlocks(paragraphXml, 'r').map(runText).join('');
    const blankParagraph = inspectBlankParagraph(paragraphXml, paragraphText);
    const leadingBreakCount = leadingManualLineBreakCount(paragraphXml);
    const hasPageBreakBefore = effective.pageBreakBefore === true;
    const fieldTextSpans = parseFieldTextSpans(paragraphXml);
    const layoutSignals: ParagraphLayoutSignals = {
      hasDrawing: /<(?:[A-Za-z_][\w.-]*:)?(?:drawing|pict|object)\b/i.test(paragraphXml),
      hasAnchoredDrawing: /<(?:[A-Za-z_][\w.-]*:)?anchor\b/i.test(paragraphXml),
      hasCustomTabs: Boolean(effective.tabStops?.length),
      hasNumbering: /<(?:[A-Za-z_][\w.-]*:)?numPr\b/i.test(paragraphXml)
        || /(?:list|bullet|number)/i.test(styleId || '')
    };

    const aggregateRunProperty = <T extends 'fontName' | 'fontSize' | 'bold' | 'italic'>(key: T): string | number | boolean | null => {
      if (!effectiveRuns.length) return effective[key];
      const values = effectiveRuns.map(run => run[key]);
      if (values.some(value => value === null)) return null;
      return values.every(value => value === values[0]) ? values[0] : null;
    };

    return {
      index,
      text: paragraphText,
      lineSpacingMultiple: effective.lineSpacingMultiple,
      lineSpacingRule: effective.lineSpacingRule,
      lineSpacingValue: effective.lineSpacingValue,
      lineSpacingSource: source,
      fontName: aggregateRunProperty('fontName') as string | null,
      fontSize: aggregateRunProperty('fontSize') as number | null,
      bold: aggregateRunProperty('bold') as boolean | null,
      italic: aggregateRunProperty('italic') as boolean | null,
      allCaps: paragraphVisualAllCaps(effectiveRuns),
      figurePrefixAllCaps: captionPrefixVisualAllCaps(effectiveRuns, 'figure'),
      tablePrefixAllCaps: captionPrefixVisualAllCaps(effectiveRuns, 'table'),
      alignment: effective.alignment,
      firstLineIndent: effective.firstLineIndent,
      leftIndent: effective.leftIndent,
      rightIndent: effective.rightIndent,
      spaceBefore: effective.spaceBefore,
      spaceAfter: effective.spaceAfter,
      formattingSources: Object.fromEntries(
        Object.entries(effective.sources).map(([key, value]) => [key, value as FormattingSource])
      ),
      effectiveRuns: effectiveRuns.map(run => ({
        ...run,
        sources: Object.fromEntries(
          Object.entries(run.sources).map(([key, value]) => [key, value as FormattingSource])
        )
      })),
      tabStops: effective.tabStops,
      containingTableIndex: tableMetadata[index]?.containingTableIndex ?? null,
      containingTableRowIndex: tableMetadata[index]?.containingTableRowIndex ?? null,
      isBlank: blankParagraph.whitespaceOnly,
      blankParagraph,
      leadingManualLineBreakCount: leadingBreakCount,
      hasPageBreakBefore,
      hasField: blankParagraph.hasField,
      fieldTextSpans,
      layoutSignals,
      referenceStyle,
      referenceNumbering,
      layoutSensitive: layoutSignals.hasDrawing || layoutSignals.hasAnchoredDrawing
    };
  });

  return markPseudoTableLayout(parsed, tableMetadata);
}

/** Exported for regression tests and future DOCX fixture tests. */
export { parseParagraphOoxml };

function findMatchingBlockEnd(xml: string, start: number, localName: 'p' | 'tbl'): number {
  const tokenPattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>|<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}\\s*>`, 'gi');
  tokenPattern.lastIndex = start;
  let depth = 0;
  for (const match of xml.matchAll(tokenPattern)) {
    const token = match[0];
    if (/^<\//.test(token)) {
      depth -= 1;
      if (depth === 0) return (match.index || 0) + token.length;
    } else if (!/\/\s*>$/.test(token)) {
      depth += 1;
    } else if (depth === 0) {
      return (match.index || 0) + token.length;
    }
  }
  return xml.length;
}

function blockText(xml: string): string {
  return allElementBlocks(stripTextBoxContent(xml), 't')
    .map(textXml => decodeBasicXmlEntities(textXml.replace(/^<[\s\S]*?>|<\/[\s\S]*?>$/g, '')))
    .join('');
}

/** Preserve the direct w:body order needed for caption/object placement. */
function parseDocumentBlocks(documentXml: string, stylesXml = ''): DocumentBlockSnapshot[] {
  const bodyStart = documentXml.match(/<(?:[A-Za-z_][\w.-]*:)?body\b[^>]*>/i);
  if (!bodyStart || bodyStart.index === undefined) return [];
  const contentStart = bodyStart.index + bodyStart[0].length;
  const bodyEnd = documentXml.search(new RegExp(`</(?:[A-Za-z_][\w.-]*:)?body\s*>`, 'i'));
  const contentEnd = bodyEnd >= contentStart ? bodyEnd : documentXml.length;
  const content = documentXml.slice(contentStart, contentEnd);
  const paragraphBlocks = mainFlowParagraphBlocksWithOffsets(documentXml);
  const starts = /<(?:[A-Za-z_][\w.-]*:)?(p|tbl)\b[^>]*>/gi;
  const blocks: DocumentBlockSnapshot[] = [];
  const tableStyles = parseTableStyleDefinitions(stylesXml);
  let tableOrdinal = 0;
  let match: RegExpExecArray | null;
  while ((match = starts.exec(content))) {
    const kind = match[1].toLowerCase() as 'p' | 'tbl';
    const end = findMatchingBlockEnd(content, match.index, kind);
    const xml = content.slice(match.index, end);
    const absoluteStart = contentStart + match.index;
    // Use the same main-flow sequence as ParagraphOoxmlSnapshot and Office.js.
    // Paragraphs inside drawing text boxes are intentionally excluded.
    const paragraphIndex = paragraphBlocks.filter(paragraph => paragraph.start < absoluteStart).length;
    const tableIndex = kind === 'tbl' ? tableOrdinal++ : -1;
    const drawingCount = (xml.match(/<(?:[A-Za-z_][\w.-]*:)?drawing\b/gi) || []).length;
    const tableText = blockText(xml);
    const isEquationLayout = kind === 'tbl'
      && /<(?:[A-Za-z_][\w.-]*:)?(?:oMath|oMathPara)\b/i.test(xml)
      && /\(\s*\d+\s*\)/.test(tableText);
    const effectiveBorders = kind === 'tbl'
      ? resolveEffectiveTableBorders(xml, tableStyles)
      : undefined;
    const headerRule: TableRuleState = kind === 'tbl'
      ? tableHeaderRuleState(xml, effectiveBorders)
      : {
        rule: null,
        rows: null,
        vertical: null,
        extraHorizontal: null,
        captionRowIndex: null,
        headerStartRowIndex: null,
        headerEndRowIndex: null,
        topRule: null
      };
    blocks.push({
      blockIndex: blocks.length,
      kind: kind === 'p' ? 'paragraph' : 'table',
      paragraphIndex: kind === 'p' ? paragraphIndex : undefined,
      tableIndex: kind === 'tbl' ? tableIndex : undefined,
      isEquationLayout: kind === 'tbl' ? isEquationLayout : undefined,
      drawingCount,
      hasInlineDrawing: /<(?:[A-Za-z_][\w.-]*:)?inline\b/i.test(xml),
      hasAnchoredDrawing: /<(?:[A-Za-z_][\w.-]*:)?anchor\b/i.test(xml),
      tableStyleId: kind === 'tbl' ? tableStyleId(xml) : undefined,
      effectiveTableBorders: kind === 'tbl' ? effectiveBorders : undefined,
      tableTopBorder: kind === 'tbl' ? (headerRule.topRule ?? borderState(xml, 'top', effectiveBorders?.top)) : undefined,
      tableBottomBorder: kind === 'tbl' ? borderState(xml, 'bottom', effectiveBorders?.bottom) : undefined,
      tableHeaderRule: kind === 'tbl' ? headerRule.rule : undefined,
      tableHeaderRowCount: kind === 'tbl' ? headerRule.rows : undefined,
      tableCaptionRowIndex: kind === 'tbl' ? headerRule.captionRowIndex : undefined,
      tableHeaderStartRowIndex: kind === 'tbl' ? headerRule.headerStartRowIndex : undefined,
      tableHeaderEndRowIndex: kind === 'tbl' ? headerRule.headerEndRowIndex : undefined,
      tableVerticalRules: kind === 'tbl' ? headerRule.vertical : undefined,
      tableExtraHorizontalRules: kind === 'tbl' ? headerRule.extraHorizontal : undefined,
      text: tableText
    });
    starts.lastIndex = end;
  }
  return blocks;
}

export { parseDocumentBlocks };

function parseLayoutWithDom(documentXml: string): ParsedLayout {
  const document = parseXmlWithRecovery(documentXml);
  const sections = Array.from(document.getElementsByTagNameNS(W_NS, 'sectPr'));
  if (!sections.length) throw new Error('No w:sectPr section properties were found in word/document.xml.');

  const parsedSections = sections.map((section, index): SectionLayoutSnapshot => {
    const pageSize = section.getElementsByTagNameNS(W_NS, 'pgSz')[0] || null;
    const margins = section.getElementsByTagNameNS(W_NS, 'pgMar')[0] || null;
    const cols = section.getElementsByTagNameNS(W_NS, 'cols')[0] || null;
    return {
      index,
      boundarySignature: sectionBoundarySignatureFromDom(section),
      pageWidthInches: twipsToInches(attr(pageSize, 'w')),
      pageHeightInches: twipsToInches(attr(pageSize, 'h')),
      orientation: (() => {
        const explicit = attr(pageSize, 'orient')?.toLowerCase();
        if (explicit === 'landscape' || explicit === 'portrait') return explicit;
        const width = twipsToInches(attr(pageSize, 'w'));
        const height = twipsToInches(attr(pageSize, 'h'));
        return width !== undefined && height !== undefined && width > height ? 'landscape' : 'portrait';
      })(),
      marginTopInches: twipsToInches(attr(margins, 'top')),
      marginRightInches: twipsToInches(attr(margins, 'right')),
      marginBottomInches: twipsToInches(attr(margins, 'bottom')),
      marginLeftInches: twipsToInches(attr(margins, 'left')),
      columns: Number(attr(cols, 'num') || '1')
    };
  });
  const finalSection = parsedSections[parsedSections.length - 1];

  // Header/footer prohibition applies to the whole manuscript, not only the
  // final section, so scan every section for references.
  const hasHeaderReference = sections.some(s => s.getElementsByTagNameNS(W_NS, 'headerReference').length > 0);
  const hasFooterReference = sections.some(s => s.getElementsByTagNameNS(W_NS, 'footerReference').length > 0);

  return {
    sections: parsedSections,
    ...finalSection,
    hasHeaderReference,
    hasFooterReference
  };
}

/**
 * Targeted fallback for Word-on-the-web environments where browser DOMParser
 * rejects an otherwise usable OOXML part. We only parse the section-level
 * elements needed by FormaScholar and do not attempt to interpret arbitrary
 * document XML with regular expressions.
 */
function parseLayoutWithFallback(documentXml: string): ParsedLayout {
  const xml = sanitizeXml(documentXml);
  const sections = allElementBlocks(xml, 'sectPr');
  if (!sections.length) throw new Error('No sectPr section properties were found in word/document.xml.');

  const parsedSections = sections.map((section, index): SectionLayoutSnapshot => {
    const pageSizeTag = elementStartTag(section, 'pgSz');
    const marginsTag = elementStartTag(section, 'pgMar');
    const colsTag = elementStartTag(section, 'cols');
    const columnCount = Number(attrFromTag(colsTag, 'num') || '1');
    return {
      index,
      boundarySignature: removeMutableSectionProperties(section),
      pageWidthInches: twipsToInches(attrFromTag(pageSizeTag, 'w')),
      pageHeightInches: twipsToInches(attrFromTag(pageSizeTag, 'h')),
      orientation: (() => {
        const explicit = attrFromTag(pageSizeTag, 'orient')?.toLowerCase();
        if (explicit === 'landscape' || explicit === 'portrait') return explicit;
        const width = twipsToInches(attrFromTag(pageSizeTag, 'w'));
        const height = twipsToInches(attrFromTag(pageSizeTag, 'h'));
        return width !== undefined && height !== undefined && width > height ? 'landscape' : 'portrait';
      })(),
      marginTopInches: twipsToInches(attrFromTag(marginsTag, 'top')),
      marginRightInches: twipsToInches(attrFromTag(marginsTag, 'right')),
      marginBottomInches: twipsToInches(attrFromTag(marginsTag, 'bottom')),
      marginLeftInches: twipsToInches(attrFromTag(marginsTag, 'left')),
      columns: Number.isFinite(columnCount) ? columnCount : 1
    };
  });
  const finalSection = parsedSections[parsedSections.length - 1];

  return {
    sections: parsedSections,
    ...finalSection,
    hasHeaderReference: sections.some(section => countLocalElements(section, 'headerReference') > 0),
    hasFooterReference: sections.some(section => countLocalElements(section, 'footerReference') > 0)
  };
}

function normalizeSliceData(data: unknown, index: number): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }
  if (Array.isArray(data)) return Uint8Array.from(data as number[]);

  throw new Error(`Unsupported Word document slice format at slice ${index}.`);
}

export class DocxInspector {
  static async inspectCurrentDocument(): Promise<DocumentLayoutSnapshot> {
    try {
      const bytes = await this.getCompressedDocument();
      if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
        throw new Error('Word returned data that is not a valid compressed DOCX package.');
      }

      const zip = await JSZip.loadAsync(bytes);
      const documentPart = zip.file('word/document.xml');
      if (!documentPart) throw new Error('word/document.xml is missing from the current document.');

      const documentXml = await documentPart.async('string');
      const stylesXml = await zip.file('word/styles.xml')?.async('string') || '';
      const themeXml = await zip.file('word/theme/theme1.xml')?.async('string') || '';
      const numberingXml = await zip.file('word/numbering.xml')?.async('string') || '';
      const paragraphOoxml = parseParagraphOoxml(documentXml, stylesXml, themeXml, numberingXml);
      const documentBlocks = parseDocumentBlocks(documentXml, stylesXml);

      let layout: ParsedLayout;
      let parser: 'DOMParser' | 'fallback';
      try {
        layout = parseLayoutWithDom(documentXml);
        parser = 'DOMParser';
      } catch (domError) {
        try {
          layout = parseLayoutWithFallback(documentXml);
          parser = 'fallback';
        } catch (fallbackError) {
          const domMessage = domError instanceof Error ? domError.message : String(domError);
          const fallbackMessage = fallbackError instanceof Error ? fallbackError.message : String(fallbackError);
          throw new Error(`OOXML layout parsing failed. DOM: ${domMessage} Fallback: ${fallbackMessage}`);
        }
      }

      const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string') || '';
      const headerParts = referencedHeaderFooterParts(documentXml, relsXml, 'header');
      const footerParts = referencedHeaderFooterParts(documentXml, relsXml, 'footer');

      let hasHeaderContent = false;
      let hasFooterContent = false;
      let containsPageNumberField = false;

      for (const name of headerParts) {
        const xml = await zip.file(name)?.async('string');
        if (!xml) continue;
        const clean = sanitizeXml(xml);
        if (hasMeaningfulHeaderFooterContent(clean)) hasHeaderContent = true;
        if (hasPageNumberField(clean)) containsPageNumberField = true;
      }
      for (const name of footerParts) {
        const xml = await zip.file(name)?.async('string');
        if (!xml) continue;
        const clean = sanitizeXml(xml);
        if (hasMeaningfulHeaderFooterContent(clean)) hasFooterContent = true;
        if (hasPageNumberField(clean)) containsPageNumberField = true;
      }

      // Page fields can also be placed in the document body. They are valid
      // Word fields only when represented as PAGE/NUMPAGES instructions, not
      // when ordinary prose merely contains the word "page".
      if (hasPageNumberField(documentXml)) containsPageNumberField = true;

      return {
        ...layout,
        paragraphOoxml,
        documentBlocks,
        // These properties retain their original names for strategy backwards
        // compatibility, but now represent actual referenced *content* rather
        // than merely an empty header/footer relationship.
        hasHeaderReference: hasHeaderContent,
        hasFooterReference: hasFooterContent,
        containsPageNumberField,
        ooxmlAvailable: true,
        ooxmlParser: parser
      };
    } catch (error) {
      return {
        ooxmlAvailable: false,
        ooxmlError: error instanceof Error ? error.message : String(error)
      };
    }
  }

  private static getCompressedDocument(): Promise<Uint8Array> {
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 4 * 1024 * 1024 },
        result => {
          if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
            reject(new Error(result.error?.message || 'Word did not provide the compressed document.'));
            return;
          }

          const file = result.value;
          const slices: Uint8Array[] = new Array(file.sliceCount);
          let next = 0;

          const closeAndReject = (message: string) => {
            file.closeAsync(() => reject(new Error(message)));
          };

          const readNext = () => {
            if (next >= file.sliceCount) {
              const total = slices.reduce((sum, part) => sum + part.length, 0);
              const merged = new Uint8Array(total);
              let offset = 0;
              for (const part of slices) {
                merged.set(part, offset);
                offset += part.length;
              }
              file.closeAsync(() => resolve(merged));
              return;
            }

            const index = next++;
            file.getSliceAsync(index, sliceResult => {
              if (sliceResult.status !== Office.AsyncResultStatus.Succeeded || !sliceResult.value) {
                closeAndReject(sliceResult.error?.message || `Unable to read document slice ${index}.`);
                return;
              }

              try {
                slices[index] = normalizeSliceData(sliceResult.value.data, index);
              } catch (error) {
                closeAndReject(error instanceof Error ? error.message : String(error));
                return;
              }

              readNext();
            });
          };

          readNext();
        }
      );
    });
  }
}
