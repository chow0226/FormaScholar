import { DocumentSnapshot, FigureCitationCorrectionResult, FrontMatterCorrectionOutcome, FrontMatterTransitionFixTarget, ParagraphSnapshot, TextReplacement } from '../core/models';
import { DocxInspector, repairParagraphIndentOoxml } from '../ooxml/DocxInspector';
import { analyzeCrossReferences } from '../ooxml/CrossReferenceAnalyzer';
import { ApplicabilityAnalyzer } from '../core/ApplicabilityAnalyzer';
import { analyzeAffiliationEmailBlankLine, analyzeFigureSpaceAbove, analyzeFrontMatterTransitions, analyzeTableCaptionTableGaps, findFigureCaptionBodyGapIssues, isStructurallySafeFrontMatterBlank } from '../core/ValidationEngine';
import { romanNumeralToNumber } from '../core/CaptionNumberParser';
import { isTechnicalToken, paragraphTextFingerprint } from '../core/utils';
import { recognizeSectionTitle } from '../core/SectionOntology';
import { classifyFieldInstruction, findWordFields, isTableSequenceFieldInstruction, type WordFieldKind } from '../ooxml/WordFieldAnalyzer';

export function tableBoundariesToRestore(
  insideHWasPresent: boolean,
  hasEmbeddedCaptionBoundary: boolean,
  hasHeaderBodyBoundary: boolean
): { captionHeader: boolean; headerBody: boolean } {
  return {
    captionHeader: insideHWasPresent && hasEmbeddedCaptionBoundary,
    headerBody: insideHWasPresent && hasHeaderBodyBoundary
  };
}

/** Portable Office.js value for the built-in Heading 1 style. */
export const BUILT_IN_HEADING1_STYLE = 'Heading1' as Word.BuiltInStyleName;

function nullableNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

function nullableBoolean(value: unknown): boolean | null {
  return typeof value === 'boolean' ? value : null;
}

function nullableString(value: unknown): string {
  return typeof value === 'string' ? value : '';
}

/** Preserve line-break count and position while remaining tolerant of CR/LF
 * differences between Office.js and OOXML text representations. */
function paragraphBreakStructure(text: string): string {
  const normalized = text.replace(/\r\n?/g, '\n').replace(/[\v\f]/g, '\n');
  const positions: number[] = [];
  let nonBreakOffset = 0;
  for (const character of normalized) {
    if (character === '\n') positions.push(nonBreakOffset);
    else nonBreakOffset += 1;
  }
  return `${positions.length}:${positions.join(',')}`;
}

/** Collapse run-level boolean values while preserving a mixed or unavailable state. */
export function aggregateBooleanState(values: boolean[]): boolean | null {
  if (values.length === 0) return null;
  if (values.every(value => value === true)) return true;
  if (values.every(value => value === false)) return false;
  return null;
}

export { paragraphTextFingerprint } from '../core/utils';

const SMALL_TITLE_WORDS = new Set([
  'a', 'an', 'the', 'and', 'but', 'or', 'nor', 'for', 'from', 'with', 'in', 'on',
  'at', 'by', 'to', 'of', 'off', 'as', 'via', 'vs', 'versus', 'per'
]);

function capitalizeSegment(segment: string): string {
  if (!segment) return segment;
  if (isTechnicalToken(segment)) return segment;
  if (/\d/.test(segment)) return segment; // Keep YOLOv8, 5G, H2O, etc.
  if (/^[A-Z0-9]+$/.test(segment)) return segment; // Preserve ordinary acronyms such as CNN, SVM, and AI.
  return segment.charAt(0).toUpperCase() + segment.slice(1).toLowerCase();
}

/**
 * Conservative deterministic title case. It preserves acronyms, mixed-case
 * technical tokens and tokens containing digits, while lowercasing common
 * short function words unless they are the first or last lexical token.
 */
export function toSafeTitleCase(text: string): string {
  const tokens = text.split(/(\s+)/);
  const lexical = tokens
    .map((token, index) => ({ token, index }))
    .filter(x => /[A-Za-z]/.test(x.token));
  const firstLexical = lexical[0]?.index ?? -1;
  const lastLexical = lexical[lexical.length - 1]?.index ?? -1;

  return tokens.map((token, index) => {
    if (!/[A-Za-z]/.test(token)) return token;

    // Work on hyphenated/slashed compounds without touching surrounding punctuation.
    const leading = token.match(/^[^A-Za-z0-9]*/)?.[0] ?? '';
    const trailing = token.match(/[^A-Za-z0-9]*$/)?.[0] ?? '';
    const core = token.slice(leading.length, token.length - trailing.length || undefined);
    if (!core) return token;

    // Check the complete token before splitting hyphenated/slashed terms.
    // This preserves technical spellings such as scikit-learn as a unit.
    if (isTechnicalToken(core)) return `${leading}${core}${trailing}`;

    const lowerCore = core.toLowerCase();
    if (index !== firstLexical && index !== lastLexical && SMALL_TITLE_WORDS.has(lowerCore)) {
      return `${leading}${lowerCore}${trailing}`;
    }

    const converted = core
      .split(/([\-/])/)
      .map(part => part === '-' || part === '/' ? part : capitalizeSegment(part))
      .join('');
    return `${leading}${converted}${trailing}`;
  }).join('');
}

const FIGURE_CAPTION_LABEL = /^\s*((?:FIGURE|FIG\.?)\s+\d+[A-Za-z]?(?:\s*[.\-:)])?)/i;

/** Return only the leading figure label; caption body text is never included. */
export function extractFigureCaptionLabel(text: string): string | null {
  return text.match(FIGURE_CAPTION_LABEL)?.[1] ?? null;
}

function startsWithCaptionKeyword(text: string, kind: 'figure' | 'table'): boolean {
  const pattern = kind === 'figure'
    ? /^(?:FIGURE|FIG\.?)\b/i
    : /^TABLE\b/i;
  return pattern.test(text.trim());
}

function isCaptionCandidate(text: string, kind: 'figure' | 'table'): boolean {
  const keyword = kind === 'figure' ? '(?:FIGURE|FIG\.?)' : 'TABLE';
  return new RegExp(`^${keyword}\\s+(?:\\d+[A-Za-z]?|[IVXLCDM]+)\\s*[.:-]`, 'i').test(text.trim());
}

function stripReferencePrefix(text: string): string {
  return text
    .replace(/^\s*(?:\[\s*\d+\s*\]|\d+\s*[.)])\s*/, '')
    .trimStart();
}

function paragraphContentRange(paragraph: Word.Paragraph): Word.Range {
  return paragraph.getRange(Word.RangeLocation.content);
}

function formatReferenceFormattingError(stage: string, error: unknown): Error {
  const details = (error && typeof error === 'object') ? error as {
    code?: unknown;
    message?: unknown;
    debugInfo?: unknown;
    traceMessages?: unknown;
  } : {};
  const message = details.message !== undefined ? String(details.message) : String(error);
  const extras = [
    details.code !== undefined ? `code=${String(details.code)}` : '',
    details.debugInfo !== undefined ? `debugInfo=${typeof details.debugInfo === 'string' ? details.debugInfo : JSON.stringify(details.debugInfo)}` : '',
    details.traceMessages !== undefined ? `traceMessages=${typeof details.traceMessages === 'string' ? details.traceMessages : JSON.stringify(details.traceMessages)}` : ''
  ].filter(Boolean);
  return new Error(`Reference formatting failed during ${stage}: ${message}${extras.length ? ` (${extras.join('; ')})` : ''}`);
}

function referenceFormattingApproximately(value: unknown, expected: number, tolerance = 0.75): boolean {
  return typeof value === 'number' && Number.isFinite(value) && Math.abs(value - expected) <= tolerance;
}

const AIP_RAW_REFERENCE_LEFT_INDENT_PT = 21.3;
const AIP_RAW_REFERENCE_FIRST_LINE_INDENT_PT = -21.3;
const USER_REFERENCE_LEFT_INDENT_PT = 0;
const USER_REFERENCE_FIRST_LINE_INDENT_PT = -21.6;

/**
 * Word can expose a paragraph style's indentation either as the raw OOXML
 * pair (left = text position, firstLine = -hanging) or as the user-facing
 * left anchor.  Keep the representation so a compliant raw AIP pair is not
 * rewritten into a different layout.
 */
function normalizedReferenceStyleLeftIndent(
  leftIndent: unknown,
  firstLineIndent: unknown
): { value: number | null; representation: 'raw' | 'userFacing' | 'unknown' } {
  if (typeof leftIndent !== 'number' || !Number.isFinite(leftIndent)) {
    return { value: null, representation: 'unknown' };
  }
  if (typeof firstLineIndent === 'number'
    && Number.isFinite(firstLineIndent)
    && firstLineIndent < 0
    && leftIndent > 0) {
    // Word's live style API exposes the raw OOXML pair as a positive total
    // left position plus a negative hanging/first-line value.  Do not require
    // the pair to be compliant before recognizing it: malformed raw pairs
    // such as 30 / -21.3 must be repairable to 21.3 / -21.3.
    return { value: leftIndent + firstLineIndent, representation: 'raw' };
  }
  if (typeof firstLineIndent === 'number' && Number.isFinite(firstLineIndent)) {
    return { value: leftIndent, representation: 'userFacing' };
  }
  return { value: null, representation: 'unknown' };
}

export interface ParagraphTypographyOptions {
  fontName?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
}

function associationMarkers(text: string, target: 'authorAffiliation' | 'authorEmail'): string[] {
  const normalized = text.replace(/\r\n/g, '\n').replace(/[\r\v\f]/g, '\n');
  if (target === 'authorAffiliation') {
    const marker = normalized.match(/^\s*(\d+(?:\s*[,;]\s*\d+)*)/)?.[1];
    return marker ? [marker] : [];
  }

  const markers = normalized
    .split('\n')
    .flatMap(line => {
      const isEmailLine = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i.test(line) || /^\s*[A-Za-z]\)/.test(line);
      return isEmailLine ? (line.match(/^\s*([A-Za-z]\))/)?.[1] || []) : [];
    });
  return [...new Set(markers)];
}

function normalizeLogicalLineBreaks(text: string): string {
  return text.replace(/\r\n/g, '\n').replace(/[\r\v\f]/g, '\n');
}

/** Return only the text span that should be italic for an author association. */
function authorItalicBodyTexts(text: string, target: 'authorAffiliation' | 'authorEmail'): string[] {
  const normalized = normalizeLogicalLineBreaks(text);
  if (target === 'authorAffiliation') {
    const match = normalized.match(/^\s*(\d+(?:\s*[,;]\s*\d+)*)(?:\s+)?(.*)$/);
    return match?.[2].trim() ? [match[2].trim()] : [];
  }

  return normalized
    .split('\n')
    .map(line => line.match(/^\s*([A-Za-z]\))(\s*)(.*)$/)?.[3]?.trim() || '')
    .filter(Boolean);
}

function classifyOfficeField(type: string, code: string): WordFieldKind {
  const codeKind = classifyFieldInstruction(code);
  if (codeKind !== 'OTHER') return codeKind;

  const normalizedType = type.toLowerCase();
  if (normalizedType === 'page') return classifyFieldInstruction('PAGE');
  if (normalizedType === 'numpages') return classifyFieldInstruction('NUMPAGES');
  if (normalizedType === 'sectionpages') return classifyFieldInstruction('SECTIONPAGES');
  return 'OTHER';
}

export function collapseRepeatedOrdinarySpaces(text: string): string {
  return text.replace(/ {2,}/g, ' ');
}

/** Merge adjacent raised runs before searching, so fragments such as
 * "1" + "," + "c" + ")" cannot match ordinary letters in an author name. */
export function groupRaisedAssociationMarkerRuns(
  runs: NonNullable<ParagraphSnapshot['effectiveRuns']>
): Array<{ text: string; fallbackSize: number | null }> {
  const groups: Array<{ text: string; fallbackSize: number | null }> = [];
  let currentText = '';
  let fallbackSize: number | null = null;

  const flush = (): void => {
    const text = currentText.replace(/\s+/g, ' ').trim();
    if (/^(?:\d+|[A-Za-z]\))(?:(?:\s*[,;]\s*)(?:\d+|[A-Za-z]\)))*$/.test(text)) {
      groups.push({ text, fallbackSize });
    }
    currentText = '';
    fallbackSize = null;
  };

  for (const run of runs) {
    const raised = run.verticalAlign === 'superscript'
      || (run.baselinePositionPt ?? 0) > 0;
    if (!raised) {
      flush();
      continue;
    }
    currentText += run.text;
    if (fallbackSize === null && run.fontSize !== null && run.fontSize !== undefined) {
      fallbackSize = run.fontSize;
    }
  }
  flush();
  return groups.filter((group, index, values) => values.findIndex(value => value.text === group.text) === index);
}

export function extractCaptionPrefix(
  text: string,
  kind: 'abstract' | 'figure' | 'table'
): string | null {
  const pattern = kind === 'abstract'
    ? /^\s*(Abstract\s*\.)/i
    : kind === 'figure'
      ? /^\s*(FIGURE\s+(?:\d+[A-Z]?|[IVXLCDM]+)\s*\.)/i
      : /^\s*(TABLE\s+(?:\d+[A-Z]?|[IVXLCDM]+)\s*\.)/i;
  return text.match(pattern)?.[1] ?? null;
}

export function hasExplicitCaptionPrefix(text: string, kind: 'figure' | 'table'): boolean {
  const pattern = kind === 'figure'
    ? /^FIGURE\s+\d+[A-Z]?\s*[.:-]/i
    : /^TABLE\s+\d+[A-Z]?\s*[.:-]/i;
  return pattern.test(text.trim());
}

export class WordDocumentAdapter {
  static async snapshot(): Promise<DocumentSnapshot> {
    const paragraphs = await Word.run(async (context) => {
      const collection = context.document.body.paragraphs;
      collection.load('items');
      await context.sync();

      for (const p of collection.items) {
        p.load('text,style,styleBuiltIn,alignment,firstLineIndent,leftIndent,rightIndent,spaceBefore,spaceAfter,lineSpacing,isListItem,tableNestingLevel');
        // Keep the Office.js request on the cross-platform WordApi surface.
        // All Caps is resolved from OOXML below because Font.allCaps is a
        // desktop-only API member and is not guaranteed on Word for the web.
        p.font.load('name,size,bold,italic');
      }
      await context.sync();

      for (const p of collection.items) {
        if (p.isListItem) p.listItem.load('listString');
      }
      await context.sync();

      const preliminary = collection.items.map((p, index) => ({
        index,
        text: nullableString(p.text),
        style: nullableString(p.style),
        styleBuiltIn: nullableString(p.styleBuiltIn),
        alignment: nullableString(p.alignment),
        firstLineIndent: nullableNumber(p.firstLineIndent),
        leftIndent: nullableNumber(p.leftIndent),
        rightIndent: nullableNumber(p.rightIndent),
        spaceBefore: nullableNumber(p.spaceBefore),
        spaceAfter: nullableNumber(p.spaceAfter),
        lineSpacing: nullableNumber(p.lineSpacing),
        // Semantic line spacing is resolved from OOXML below. Office.js
        // exposes points but does not expose the line-spacing rule or style
        // inheritance needed to distinguish single, exact and atLeast.
        lineSpacingMultiple: null,
        lineSpacingRule: null as string | null,
        lineSpacingValue: null as number | null,
        tableNestingLevel: nullableNumber(p.tableNestingLevel) ?? 0,
        fontName: typeof p.font.name === 'string' ? p.font.name : null,
        fontSize: nullableNumber(p.font.size),
        bold: nullableBoolean(p.font.bold),
        italic: nullableBoolean(p.font.italic),
        allCaps: null,
        // Word normally exposes automatic lists through isListItem, but some
        // imported/template paragraphs expose only a list/bulleted style.
        // Preserve that signal so normal body-indent rules do not overwrite a
        // hanging list layout.
        isListItem: Boolean(p.isListItem)
          || /(?:\blist\b|bulleted|numbered)/i.test(`${p.style} ${p.styleBuiltIn}`),
        listString: p.isListItem ? nullableString(p.listItem.listString) || null : null,
        abstractPrefixBold: null as boolean | null,
        abstractBodyBold: null as boolean | null,
        abstractBodyMixed: null as boolean | null,
        figurePrefixBold: null as boolean | null,
        figurePrefixAllCaps: null as boolean | null,
        tablePrefixBold: null as boolean | null,
        tablePrefixAllCaps: null as boolean | null
      }));

      // Mixed-format rules cannot be inferred from paragraph.font because Word
      // returns null when only part of a paragraph is bold. Inspect the exact
      // required label/prefix range instead.
      const prefixSearches: Array<{
        index: number;
        kind: 'abstract' | 'figure' | 'table';
        ranges: Word.RangeCollection;
        bodyRange?: Word.Range;
      }> = [];

      for (const item of preliminary) {
        let prefix: string | null = null;
        let kind: 'abstract' | 'figure' | 'table' | null = null;
        const abstract = extractCaptionPrefix(item.text, 'abstract');
        // Prefix formatting is independent from capitalization validation: a
        // correctly bold "Figure 1." should pass the bold rule while still
        // failing the separate uppercase-label rule.
        const figure = extractCaptionPrefix(item.text, 'figure');
        const table = extractCaptionPrefix(item.text, 'table');
        if (abstract) { prefix = abstract; kind = 'abstract'; }
        else if (figure) { prefix = figure; kind = 'figure'; }
        else if (table) { prefix = table; kind = 'table'; }
        if (!prefix || !kind) continue;

        const ranges = collection.items[item.index].getRange().search(prefix, {
          matchCase: true,
          matchWholeWord: false
        });
        ranges.load('items');
        prefixSearches.push({ index: item.index, kind, ranges });
      }

      if (prefixSearches.length) {
        await context.sync();
        for (const search of prefixSearches) {
          const first = search.ranges.items[0];
          if (first) {
            first.font.load('bold');
            if (search.kind === 'abstract') {
              // Build a real range from the end of the exact prefix to the
              // end of the paragraph. RangeLocation.After is only a point.
              const prefixEnd = first.getRange('End');
              const paragraphEnd = collection.items[search.index].getRange('End');
              search.bodyRange = prefixEnd.expandTo(paragraphEnd);
              search.bodyRange.font.load('bold');
            }
          }
        }
        await context.sync();

        for (const search of prefixSearches) {
          const first = search.ranges.items[0];
          const value = first ? nullableBoolean(first.font.bold) : null;
          const item = preliminary[search.index];
          if (search.kind === 'abstract') {
            item.abstractPrefixBold = value;
            if (search.bodyRange) {
              const bodyBold = nullableBoolean(search.bodyRange.font.bold);
              item.abstractBodyBold = bodyBold;
              item.abstractBodyMixed = bodyBold === null;
            }
          }
          else if (search.kind === 'figure') {
            item.figurePrefixBold = value;
          }
          else {
            item.tablePrefixBold = value;
          }
        }
      }

      const classified = this.classify(preliminary);
      const authorBodySearches: Array<{
        index: number;
        ranges: Word.RangeCollection;
      }> = [];

      // Whole author paragraphs are intentionally mixed: association markers
      // are superscript/non-italic while the affiliation or email body is
      // italic. Read the exact live body ranges so a just-applied correction
      // is not hidden by stale compressed-DOCX run data.
      for (const item of classified) {
        const target = item.classification === 'authorEmail'
          ? 'authorEmail'
          : item.classification === 'authorAffiliation'
            ? 'authorAffiliation'
            : null;
        if (!target) continue;

        for (const bodyText of authorItalicBodyTexts(item.text, target)) {
          const ranges = collection.items[item.index]
            .getRange(Word.RangeLocation.content)
            .search(bodyText, {
              matchCase: true,
              matchWholeWord: false
            });
          ranges.load('items');
          authorBodySearches.push({ index: item.index, ranges });
        }
      }

      if (authorBodySearches.length) {
        await context.sync();
        for (const search of authorBodySearches) {
          for (const range of search.ranges.items) range.font.load('italic');
        }
        await context.sync();

        const valuesByParagraph = new Map<number, boolean[]>();
        for (const search of authorBodySearches) {
          const values = valuesByParagraph.get(search.index) || [];
          for (const range of search.ranges.items) {
            const value = nullableBoolean(range.font.italic);
            if (value !== null) values.push(value);
          }
          valuesByParagraph.set(search.index, values);
        }

        for (const item of classified) {
          if (item.classification !== 'authorEmail'
            && item.classification !== 'authorAffiliation') continue;
          const values = valuesByParagraph.get(item.index) || [];
          item.authorAssociationBodyItalic = aggregateBooleanState(values);
        }
      }

      return classified;
    });

    const abstractIndex = paragraphs.findIndex(p => p.classification === 'abstract');
    const referencesHeadingIndex = paragraphs.findIndex(p => p.classification === 'referencesHeading');
    const layout = await DocxInspector.inspectCurrentDocument();
    const ooxmlParagraphList = layout.paragraphOoxml || [];
    const ooxmlParagraphs = new Map(ooxmlParagraphList.map(p => [p.index, p]));
    const paragraphCountAligned = ooxmlParagraphList.length === paragraphs.length;
    const enrichedParagraphs = paragraphs.map(p => {
      // A complex drawing can make one OOXML paragraph differ from its live
      // Office.js counterpart. Do not disable enrichment for the whole
      // document; merge only when this paragraph's text agrees exactly.
      const candidate = paragraphCountAligned ? ooxmlParagraphs.get(p.index) : undefined;
      const textAligned = Boolean(candidate)
        && paragraphTextFingerprint(p.text) === paragraphTextFingerprint(candidate?.text || '');
      if (!candidate || !textAligned) {
        return {
          ...p,
          ooxmlAligned: false,
          leadingManualLineBreakCount: null
        };
      }
      const ooxml = candidate;
      return {
        ...p,
        ooxmlAligned: true,
        lineSpacingMultiple: ooxml.lineSpacingMultiple,
        lineSpacingRule: ooxml.lineSpacingRule,
        lineSpacingValue: ooxml.lineSpacingValue,
        // Office.js reflects the current live Word range after a correction.
        // Prefer it whenever the paragraph has a uniform value; use OOXML
        // only when Word reports null for mixed/unavailable formatting.
        fontName: p.fontName !== null ? p.fontName : ooxml.fontName ?? null,
        fontSize: p.fontSize !== null ? p.fontSize : ooxml.fontSize ?? null,
        bold: p.bold !== null ? p.bold : ooxml.bold ?? null,
        italic: p.italic !== null ? p.italic : ooxml.italic ?? null,
        allCaps: ooxml.allCaps !== undefined ? ooxml.allCaps : p.allCaps,
        figurePrefixAllCaps: ooxml.figurePrefixAllCaps !== undefined ? ooxml.figurePrefixAllCaps : p.figurePrefixAllCaps,
        tablePrefixAllCaps: ooxml.tablePrefixAllCaps !== undefined ? ooxml.tablePrefixAllCaps : p.tablePrefixAllCaps,
        alignment: ooxml.alignment !== undefined && ooxml.alignment !== null ? ooxml.alignment : p.alignment,
        firstLineIndent: ooxml.firstLineIndent !== undefined ? ooxml.firstLineIndent : p.firstLineIndent,
        leftIndent: ooxml.leftIndent !== undefined ? ooxml.leftIndent : p.leftIndent,
        rightIndent: ooxml.rightIndent !== undefined ? ooxml.rightIndent : p.rightIndent,
        spaceBefore: ooxml.spaceBefore !== undefined ? ooxml.spaceBefore : p.spaceBefore,
        spaceAfter: ooxml.spaceAfter !== undefined ? ooxml.spaceAfter : p.spaceAfter,
        tabStops: ooxml.tabStops,
        containingTableIndex: ooxml.containingTableIndex,
        containingTableRowIndex: ooxml.containingTableRowIndex,
        isBlank: ooxml.isBlank,
        blankParagraph: ooxml.blankParagraph,
        leadingManualLineBreakCount: ooxml.leadingManualLineBreakCount,
        hasPageBreakBefore: ooxml.hasPageBreakBefore,
        hasField: ooxml.hasField === true,
        fieldTextSpans: ooxml.fieldTextSpans,
        layoutSensitive: ooxml.layoutSensitive === true,
        layoutSignals: ooxml.layoutSignals,
        referenceStyle: ooxml.referenceStyle,
        referenceNumbering: (() => {
          const manualPrefix = /^\s*(?:\[\s*\d+\s*\]|\d+\s*[.)])\s+/.test(p.text);
          const parsed = ooxml.referenceNumbering || {
            kind: 'none' as const,
            level: null,
            format: null,
            suffix: null,
            startAt: null,
            numberPositionPt: null,
            textPositionPt: null
          };
          const automatic = parsed.kind === 'automatic' || p.isListItem;
          return {
            ...parsed,
            kind: automatic && manualPrefix
              ? 'mixed'
              : automatic
                ? 'automatic'
                : manualPrefix
                  ? 'manual'
                  : parsed.kind
          };
        })(),
        formattingSources: ooxml.formattingSources,
        effectiveRuns: ooxml.effectiveRuns
      };
    });

    // OOXML-only layout signals are available only after the second pass.
    // Reclassify once so pseudo-table rows cannot fall through to body merely
    // because their imported style says Normal/Body Text.
    const finalParagraphs = this.classify(enrichedParagraphs);
    const snapshot: DocumentSnapshot = {
      paragraphs: finalParagraphs,
      layout: {
        ...layout,
        paragraphMappingAligned: paragraphCountAligned
      },
      crossReferences: analyzeCrossReferences(
        finalParagraphs,
        layout.documentBlocks || [],
        finalParagraphs.find(p => p.classification === 'referencesHeading')?.index ?? null
      ),
      abstractIndex: finalParagraphs.findIndex(p => p.classification === 'abstract') >= 0
        ? finalParagraphs.findIndex(p => p.classification === 'abstract')
        : null,
      referencesHeadingIndex: finalParagraphs.findIndex(p => p.classification === 'referencesHeading') >= 0
        ? finalParagraphs.findIndex(p => p.classification === 'referencesHeading')
        : null
    };
    snapshot.features = ApplicabilityAnalyzer.analyze(snapshot);
    return snapshot;
  }

  static classify(input: Omit<ParagraphSnapshot, 'classification'>[]): ParagraphSnapshot[] {
    const firstNonBlank = input.find(p => p.text.trim().length > 0)?.index ?? -1;
    const abstractIndex = input.find(p => {
      const style = `${p.style} ${p.styleBuiltIn}`.toLowerCase();
      return /abstract/.test(style) || /^\s*Abstract\b/i.test(p.text);
    })?.index ?? -1;
    const referenceHeadingCandidates = input.filter(p => {
      const t = p.text.trim();
      const style = `${p.style} ${p.styleBuiltIn}`.toLowerCase();
      return p.tableNestingLevel === 0
        && (abstractIndex < 0 || p.index > abstractIndex)
        && /^(?:references|bibliography)\s*$/i.test(t)
        && (/heading\s*1/.test(style) || /^(?:references|bibliography)\s*$/i.test(t));
    });
    const styledReferenceHeading = referenceHeadingCandidates.find(p => {
      const style = `${p.style} ${p.styleBuiltIn}`.toLowerCase();
      return /heading\s*1/.test(style);
    });
    const referencesIndex = styledReferenceHeading?.index
      ?? referenceHeadingCandidates[referenceHeadingCandidates.length - 1]?.index
      ?? -1;

    // Reference entries are a contiguous numbered/reference-styled block after
    // the REFERENCES heading. This prevents later appendices/template guidance
    // from being misclassified as references merely because they occur later.
    const isLikelyUnnumberedReference = (text: string): boolean => {
      const normalized = text.trim();
      if (normalized.length < 20 || normalized.split(/\s+/).length < 5) return false;
      const commaCount = (normalized.match(/,/g) || []).length;
      const bibliographicPunctuation = (normalized.match(/[,.]/g) || []).length;
      const hasDoi = /\bdoi\s*:\s*|\bdoi\.org\//i.test(normalized);
      const hasYear = /\b(?:18|19|20|21)\d{2}[a-z]?\b/i.test(normalized);
      const hasPageRange = /\b\d+\s*[\-–—]\s*\d+\b/.test(normalized);
      return hasDoi
        || (hasYear && commaCount >= 1)
        || (hasPageRange && commaCount >= 1)
        || bibliographicPunctuation >= 3;
    };
    const referenceIndices = new Set<number>();
    if (referencesIndex >= 0) {
      let started = false;
      for (const p of input.filter(x => x.index > referencesIndex)) {
        const t = p.text.trim();
        const style = `${p.style} ${p.styleBuiltIn}`.toLowerCase();
        const listNumber = (p.listString || '').trim();
        if (p.tableNestingLevel > 0) continue;
        const looksReference = /(^|\s)reference($|\s)/.test(style)
          || /^\s*(?:\[\s*\d+\s*\]|\d+\s*[.)])\s*/.test(t)
          || (p.isListItem && /^\d+[.)]?$/.test(listNumber))
          || p.referenceNumbering?.kind === 'automatic'
          || isLikelyUnnumberedReference(t);

        if (!started) {
          if (!t) continue;
          if (looksReference) {
            started = true;
            referenceIndices.add(p.index);
          }
          continue;
        }

        if (!t) continue;
        if (!looksReference) break;
        referenceIndices.add(p.index);
      }
    }

    return input.map((p): ParagraphSnapshot => {
      const t = p.text.trim();
      const style = `${p.style} ${p.styleBuiltIn}`.toLowerCase();
      let classification: ParagraphSnapshot['classification'] = 'other';

      if (!t) classification = 'other';
      else if (p.index === firstNonBlank) classification = 'title';
      else if (p.index === referencesIndex) classification = 'referencesHeading';
      else if (referenceIndices.has(p.index)) classification = 'reference';
      // Explicit semantic prefixes outrank style names. Imported Word
      // documents often assign a Figure Caption style to a TABLE caption
      // (or vice versa); the visible TABLE/FIGURE label is the safer signal.
      else if (isCaptionCandidate(t, 'figure')) classification = 'figureCaption';
      else if (isCaptionCandidate(t, 'table')) classification = 'tableCaption';
      // Keep malformed but explicitly labelled captions with their semantic
      // kind when a caption style is present. A bad TABLE label must never be
      // treated as a figure merely because its imported style says Figure.
      else if (startsWithCaptionKeyword(t, 'table') && /caption/.test(style)) classification = 'tableCaption';
      else if (startsWithCaptionKeyword(t, 'figure') && /caption/.test(style)) classification = 'figureCaption';
      // Positioned/drawing-based paragraphs are not ordinary manuscript prose
      // even when their imported style says Normal or Body Text.
      else if (p.layoutSensitive === true) classification = 'other';
      else if (/figure\s*caption/.test(style)) classification = 'figureCaption';
      else if (/table\s*caption/.test(style)) classification = 'tableCaption';
      else if (p.tableNestingLevel > 0) classification = 'other';
      else if (/heading\s*1/.test(style)) classification = 'heading1';
      else if (/heading\s*2/.test(style)) classification = 'heading2';
      else if (/heading\s*3/.test(style)) classification = 'heading3';
      else if (/equation/.test(style)) classification = 'equation';
      else if (p.index === abstractIndex || /^\s*Abstract\s*\./i.test(t)) classification = 'abstract';
      else if (abstractIndex > 0 && p.index < abstractIndex) {
        if (/\S+@\S+\.\S+/.test(t)) classification = 'authorEmail';
        else if (/author\s*email/.test(style)) classification = 'authorEmail';
        else if (/author\s*name/.test(style)) classification = 'authorName';
        else if (/author\s*affiliation/.test(style)) classification = 'authorAffiliation';
        else if (/affiliation|university|institute|department|faculty|college|school|laboratory|centre|center/i.test(t)) classification = 'authorAffiliation';
        else classification = 'authorName';
      } else if (/^\s*Abstract\s*\./i.test(t)) classification = 'abstract';
      else if (/paragraph/.test(style) || /normal/.test(style) || (abstractIndex >= 0 && p.index > abstractIndex && (referencesIndex < 0 || p.index < referencesIndex))) classification = 'body';

      const sectionRecognition = recognizeSectionTitle(t, {
        classification,
        style,
        alignment: p.alignment,
        bold: p.bold,
        allCaps: p.allCaps,
        fontSize: p.fontSize,
        isListItem: p.isListItem,
        tableNestingLevel: p.tableNestingLevel,
        layoutSensitive: p.layoutSensitive
      });

      return {
        ...p,
        classification,
        sectionFamilies: sectionRecognition.canonicalSections,
        sectionNormalizedTitle: sectionRecognition.normalizedTitle,
        sectionConfidence: sectionRecognition.confidence,
        heading1Confidence: sectionRecognition.heading1Confidence
      };
    });
  }

  static async locateParagraph(index: number, expectedText?: string): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();
      const paragraph = paragraphs.items[index];
      if (!paragraph) throw new Error(`Paragraph ${index + 1} is no longer available. Revalidate the document.`);
      if (expectedText !== undefined) {
        paragraph.load('text');
        await context.sync();
        if (paragraphTextFingerprint(paragraph.text || '') !== paragraphTextFingerprint(expectedText)) {
          throw new Error(`Paragraph ${index + 1} no longer contains the text that was validated. Revalidate the document.`);
        }
      }
      paragraph.getRange().select();
      await context.sync();
    });
  }

  /** Return unique paragraph indices in the safe bottom-to-top deletion order. */
  static sortParagraphIndicesForDeletion(indices: number[]): number[] {
    return [...new Set(indices)]
      .filter(index => Number.isInteger(index) && index >= 0)
      .sort((a, b) => b - a);
  }

  /** Delete complete empty paragraph nodes, preserving all paragraph content around them. */
  static async removeBlankParagraphs(indices: number[]): Promise<void> {
    const orderedIndices = this.sortParagraphIndicesForDeletion(indices);
    if (!orderedIndices.length) return;

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      for (const index of orderedIndices) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) {
          throw new Error(
            `Blank paragraph ${index + 1} is no longer available. Revalidate the document.`
          );
        }
        paragraph.delete();
      }
      await context.sync();
    });
  }

  /** Delete only paragraphs that are still safe, empty front-matter separators. */
  static async removeSafeBlankParagraphs(indices: number[]): Promise<void> {
    const requested = [...new Set(indices)];
    if (!requested.length) return;

    const current = await this.snapshot();
    const byIndex = new Map(current.paragraphs.map(paragraph => [paragraph.index, paragraph]));
    const unsafe = requested.filter(index => {
      const paragraph = byIndex.get(index);
      return !paragraph || !isStructurallySafeFrontMatterBlank(paragraph);
    });
    if (unsafe.length) {
      throw new Error(
        `Front-matter blank paragraph mapping is no longer safe for ${unsafe.length} paragraph(s). Revalidate the document.`
      );
    }

    await this.removeBlankParagraphs(requested);
  }

  /** Normalize applicable figure-caption/body gaps to exactly one blank paragraph. */
  static async applyFigureCaptionBodyGaps(): Promise<void> {
    let current = await this.snapshot();
    let issues = findFigureCaptionBodyGapIssues(current.paragraphs);
    const extraBlankIndices = [...new Set(
      issues
        .filter(issue => !issue.blockedByStructuralParagraph)
        .flatMap(issue => issue.blankParagraphIndices.slice(1))
    )];

    // Remove surplus separators first so all later insertion indices are fresh.
    if (extraBlankIndices.length) {
      await this.removeBlankParagraphs(extraBlankIndices);
      current = await this.snapshot();
      issues = findFigureCaptionBodyGapIssues(current.paragraphs);
    }

    const missingCaptionIndices = issues
      .filter(issue => !issue.blockedByStructuralParagraph && issue.blankParagraphIndices.length === 0)
      .map(issue => issue.captionParagraphIndex)
      .sort((a, b) => b - a);
    if (!missingCaptionIndices.length) return;

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      for (const index of missingCaptionIndices) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) {
          throw new Error(`Figure caption paragraph ${index + 1} is no longer available. Revalidate the document.`);
        }
        paragraph.insertParagraph('', Word.InsertLocation.after);
      }
      await context.sync();
    });
  }

  /** Normalize the safe subset of figure spacing above to one blank paragraph. */
  static async applyFigureSpaceAbove(): Promise<void> {
    let current = await this.snapshot();
    let analysis = analyzeFigureSpaceAbove(current);
    const extraBlankIndices = [...new Set(
      analysis.issues
        .filter(issue => !issue.blockedByStructuralParagraph)
        .flatMap(issue => issue.blankParagraphIndices.slice(1))
    )];

    if (extraBlankIndices.length) {
      await this.removeBlankParagraphs(extraBlankIndices);
      current = await this.snapshot();
      analysis = analyzeFigureSpaceAbove(current);
    }

    const missingAnchors = analysis.issues
      .filter(issue => !issue.blockedByStructuralParagraph && issue.blankParagraphIndices.length === 0)
      .map(issue => issue.previousParagraphIndex)
      .sort((a, b) => b - a);
    if (!missingAnchors.length) return;

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      for (const index of [...new Set(missingAnchors)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) {
          throw new Error(`Figure spacing anchor paragraph ${index + 1} is no longer available. Revalidate the document.`);
        }
        paragraph.insertParagraph('', Word.InsertLocation.after);
      }
      await context.sync();
    });
  }

  /** Remove only safely removable blank paragraphs between external table
   * captions and their high-confidence associated tables. */
  static async applyTableCaptionTableGaps(): Promise<void> {
    const current = await this.snapshot();
    const analysis = analyzeTableCaptionTableGaps(current);
    const blankIndices = [...new Set(
      analysis.issues
        .filter(issue => !issue.blockedByStructuralParagraph)
        .flatMap(issue => issue.blankParagraphIndices)
    )];

    if (blankIndices.length) await this.removeBlankParagraphs(blankIndices);
  }

  /** Enforce one effective affiliation/email separator without duplicating
   * an existing leading manual line break in the email paragraph. */
  static async applyAffiliationEmailBlankLine(): Promise<void> {
    let current = await this.snapshot();
    let analysis = analyzeAffiliationEmailBlankLine(current.paragraphs);
    if (analysis.state === 'not_applicable' || analysis.state === 'passed') return;
    if (analysis.state === 'blocked') {
      throw new Error(analysis.reason || 'The affiliation-email separator contains protected structural content.');
    }

    const emailSignatures = current.paragraphs
      .filter(paragraph => paragraph.classification === 'authorEmail')
      .map(paragraph => ({
        text: paragraphTextFingerprint(paragraph.text),
        fields: JSON.stringify(paragraph.fieldTextSpans || []),
        breaks: paragraphBreakStructure(paragraph.text)
      }));
    const firstEmail = analysis.firstEmailParagraphIndex === undefined
      ? undefined
      : current.paragraphs.find(paragraph => paragraph.index === analysis.firstEmailParagraphIndex);
    const firstEmailFingerprint = firstEmail
      ? paragraphTextFingerprint(firstEmail.text)
      : null;
    const initialLeadingBreakCount = analysis.leadingManualBreakCount;

    if (analysis.state === 'extra' && analysis.removableBlankParagraphIndices.length) {
      await this.removeSafeBlankParagraphs(analysis.removableBlankParagraphIndices);
      current = await this.snapshot();
      analysis = analyzeAffiliationEmailBlankLine(current.paragraphs);
    }

    if (analysis.state === 'missing') {
      const anchorIndex = analysis.lastAffiliationParagraphIndex;
      if (anchorIndex === undefined) {
        throw new Error('The last affiliation paragraph is no longer available. Revalidate the document.');
      }
      await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        const paragraph = paragraphs.items[anchorIndex];
        if (!paragraph) {
          throw new Error(`Affiliation paragraph ${anchorIndex + 1} is no longer available. Revalidate the document.`);
        }
        paragraph.insertParagraph('', Word.InsertLocation.after);
        await context.sync();
      });
    }

    const after = await this.snapshot();
    const verified = analyzeAffiliationEmailBlankLine(after.paragraphs);
    if (verified.state !== 'passed') {
      throw new Error(`Affiliation-email separator verification found state '${verified.state}'.`);
    }

    const afterEmailSignatures = after.paragraphs
      .filter(paragraph => paragraph.classification === 'authorEmail')
      .map(paragraph => ({
        text: paragraphTextFingerprint(paragraph.text),
        fields: JSON.stringify(paragraph.fieldTextSpans || []),
        breaks: paragraphBreakStructure(paragraph.text)
      }));
    const contentPreserved = emailSignatures.length === afterEmailSignatures.length
      && emailSignatures.every((signature, index) =>
        signature.text === afterEmailSignatures[index].text
        && signature.fields === afterEmailSignatures[index].fields
        && signature.breaks === afterEmailSignatures[index].breaks
      );
    if (!contentPreserved) {
      throw new Error('Affiliation-email separator correction changed email text or field-backed content; manual review is required.');
    }

    if (firstEmailFingerprint !== null && initialLeadingBreakCount !== null && initialLeadingBreakCount !== undefined) {
      const matchingEmail = after.paragraphs.find(paragraph =>
        paragraph.classification === 'authorEmail'
        && paragraphTextFingerprint(paragraph.text) === firstEmailFingerprint
      );
      if (!matchingEmail || matchingEmail.leadingManualLineBreakCount !== initialLeadingBreakCount) {
        throw new Error('Affiliation-email separator correction changed the email paragraph line-break structure; manual review is required.');
      }
    }
  }

  /** Remove safe empty paragraphs from front-matter transitions that require no separator. */
  static async applyFrontMatterExtraBlankLines(
    approvedTargets: FrontMatterTransitionFixTarget[] = []
  ): Promise<FrontMatterCorrectionOutcome> {
    if (!approvedTargets.length) {
      throw new Error('Front-matter correction did not include validated transition targets. Revalidate the document.');
    }

    const current = await this.snapshot();
    const transitions = analyzeFrontMatterTransitions(current.paragraphs);
    const safeIndices: number[] = [];
    let alreadyResolvedTargets = 0;

    for (const target of approvedTargets) {
      const baseMatch = transitions.find(transition =>
        transition.fromType === target.fromType
        && transition.toType === target.toType
        && transition.fromIndex === target.fromIndex
        && transition.toIndex === target.toIndex
        && transition.fromTextFingerprint === target.fromTextFingerprint
        && transition.toTextFingerprint === target.toTextFingerprint
      );
      if (!baseMatch) {
        throw new Error('A validated front-matter transition no longer matches the document. Revalidate the document.');
      }

      if (baseMatch.blankIndices.length === 0) {
        alreadyResolvedTargets += 1;
        continue;
      }
      if (baseMatch.expectedBlankCount !== target.expectedBlankCount
        || target.expectedBlankCount !== 0
        || baseMatch.interveningNonBlankIndices.length > 0
        || baseMatch.unsafeBlankIndices.length > 0
        || baseMatch.blankIndices.length !== target.blankIndices.length
        || baseMatch.blankIndices.some((index, position) => index !== target.blankIndices[position])) {
        throw new Error('A validated front-matter transition changed or contains protected content. Revalidate the document.');
      }
      safeIndices.push(...baseMatch.safeBlankIndices);
    }

    const uniqueSafeIndices = [...new Set(safeIndices)];
    if (uniqueSafeIndices.length) {
      await this.removeSafeBlankParagraphs(uniqueSafeIndices);
    } else if (!alreadyResolvedTargets) {
      throw new Error('No safely removable front-matter blank paragraph remains. Revalidate the document.');
    }

    const verified = analyzeFrontMatterTransitions((await this.snapshot()).paragraphs);
    const consumed = new Set<number>();
    for (const target of approvedTargets) {
      const matchIndex = verified.findIndex((transition, index) => {
        if (consumed.has(index)) return false;
        return transition.fromType === target.fromType
          && transition.toType === target.toType
          && transition.fromTextFingerprint === target.fromTextFingerprint
          && transition.toTextFingerprint === target.toTextFingerprint
          && transition.expectedBlankCount === 0
          && transition.interveningNonBlankIndices.length === 0
          && transition.blankIndices.length === 0;
      });
      if (matchIndex < 0) {
        throw new Error('Front-matter blank correction verification found an approved transition that was not resolved.');
      }
      consumed.add(matchIndex);
    }

    return { removedBlankIndices: uniqueSafeIndices, alreadyResolvedTargets };
  }

  /** Locate the full citation, then replace only its approved Figure/Fig. prefix. */
  static async normalizeFigureCitationSpelling(replacements: TextReplacement[]): Promise<FigureCitationCorrectionResult> {
    if (!replacements.length) return { applied: 0, skippedFieldBacked: [] };

    return Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      const searches = replacements.map(replacement => {
        const paragraph = paragraphs.items[replacement.paragraphIndex];
        if (!paragraph) {
          throw new Error(`Figure citation paragraph ${replacement.paragraphIndex + 1} is no longer available. Revalidate the document.`);
        }
        const ranges = paragraphContentRange(paragraph).search(replacement.searchText, {
          matchCase: true,
          matchWholeWord: true
        });
        ranges.load('items');
        return { replacement, ranges };
      });
      await context.sync();

      const targets = searches.map(({ replacement, ranges }) => {
        const citationRange = ranges.items[replacement.occurrenceIndex];
        if (!citationRange) {
          throw new Error(`Figure citation "${replacement.searchText}" is no longer available in paragraph ${replacement.paragraphIndex + 1}. Revalidate the document.`);
        }

        const citationOoxml = citationRange.getOoxml();
        const prefixRanges = citationRange.search(replacement.originalPrefix, {
          matchCase: true,
          matchWholeWord: true
        });
        prefixRanges.load('items');
        return { replacement, citationOoxml, prefixRanges };
      });
      await context.sync();

      const safeTargets: Array<{
        replacement: TextReplacement;
        prefixRange: Word.Range;
      }> = [];
      const skippedFieldBacked: FigureCitationCorrectionResult['skippedFieldBacked'] = [];

      for (const { replacement, citationOoxml, prefixRanges } of targets) {
        if (findWordFields(citationOoxml.value).length > 0) {
          skippedFieldBacked.push({
            paragraphIndex: replacement.paragraphIndex,
            citation: replacement.searchText
          });
          continue;
        }
        const prefixRange = prefixRanges.items[0];
        if (!prefixRange) {
          throw new Error(`Figure citation prefix "${replacement.originalPrefix}" is no longer available in paragraph ${replacement.paragraphIndex + 1}. Revalidate the document.`);
        }
        safeTargets.push({ replacement, prefixRange });
      }

      for (const { replacement, prefixRange } of safeTargets) {
        prefixRange.insertText(replacement.replacementPrefix, Word.InsertLocation.replace);
      }
      if (safeTargets.length) await context.sync();
      return {
        applied: safeTargets.length,
        skippedFieldBacked
      };
    });
  }

  static async clearHeaders(): Promise<void> {
    await this.clearHeaderFooter('header');
  }

  static async clearFooters(): Promise<void> {
    await this.clearHeaderFooter('footer');
  }

  private static async clearHeaderFooter(kind: 'header' | 'footer'): Promise<void> {
    await Word.run(async (context) => {
      const sections = context.document.sections;
      sections.load('items');
      await context.sync();

      const types: Word.HeaderFooterType[] = [
        Word.HeaderFooterType.primary,
        Word.HeaderFooterType.firstPage,
        Word.HeaderFooterType.evenPages
      ];

      for (const section of sections.items) {
        for (const type of types) {
          const body = kind === 'header' ? section.getHeader(type) : section.getFooter(type);
          body.clear();
        }
      }

      await context.sync();
    });
  }

  static async applyTitleCase(indices: number[]): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      const targets = indices
        .map(index => paragraphs.items[index])
        .filter((p): p is Word.Paragraph => Boolean(p));
      for (const p of targets) p.load('text');
      await context.sync();

      for (const p of targets) {
        const revised = toSafeTitleCase(p.text);
        if (revised !== p.text) paragraphContentRange(p).insertText(revised, Word.InsertLocation.replace);
      }
      await context.sync();
    });
  }

  static async normalizeFigureCaptionPrefixes(failedIndices: number[], allCaptionIndices: number[]): Promise<void> {
    const failedSet = new Set(failedIndices);

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      const captions = allCaptionIndices
        .map(index => ({ index, paragraph: paragraphs.items[index] }))
        .filter((x): x is { index: number; paragraph: Word.Paragraph } => Boolean(x.paragraph));
      for (const x of captions) x.paragraph.load('text');
      await context.sync();

      const searches: Array<{
        paragraph: Word.Paragraph;
        prefix: string;
        ranges: Word.RangeCollection | null;
        insertOnly: boolean;
      }> = [];
      let figureNumber = 0;
      captions.forEach((x) => {
        const captionText = (x.paragraph.text || '').trim();
        // Defense in depth: any explicitly TABLE-prefixed paragraph is
        // protected, including malformed labels such as "TABLE I." or
        // "TABLE Results" that do not satisfy the strict caption parser.
        if (startsWithCaptionKeyword(captionText, 'table')) return;
        figureNumber += 1;
        if (!failedSet.has(x.index)) return;
        const prefix = `FIGURE ${figureNumber}.`;
        const existingPrefix = extractFigureCaptionLabel(x.paragraph.text || '');
        // A caption that already starts with FIGURE/Fig. but has an
        // unparseable label (for example "FIGURE I.") is ambiguous. Do not
        // prepend a second label; leave it for manual correction.
        if (!existingPrefix && startsWithCaptionKeyword(captionText, 'figure')) return;
        const content = paragraphContentRange(x.paragraph);
        const ranges = existingPrefix
          ? content.search(existingPrefix, { matchCase: true, matchWholeWord: false })
          : null;
        ranges?.load('items');
        searches.push({
          paragraph: x.paragraph,
          prefix,
          ranges,
          insertOnly: !existingPrefix
        });
      });
      await context.sync();

      // Make only the required FIGURE n. label bold without forcing the whole
      // caption to bold. This also prepares the paragraph for the later
      // FIGURE_CAPTION_BOLD_PREFIX rule when that rule becomes automated.
      const replaceTargets: Array<{
        range: Word.Range;
        prefix: string;
        ooxml: OfficeExtension.ClientResult<string>;
      }> = [];
      for (const item of searches) {
        if (item.insertOnly) {
          const inserted = paragraphContentRange(item.paragraph)
            .insertText(`${item.prefix} `, Word.InsertLocation.start);
          inserted.font.bold = true;
          continue;
        }

        const range = item.ranges?.items[0];
        if (!range) {
          throw new Error(`Figure caption prefix "${item.prefix}" is no longer available. Revalidate the document.`);
        }
        // A figure number can itself be a REF/cross-reference field. Do not
        // replace a range that contains field markup; the body remains safe,
        // and the caption can be reviewed manually instead.
        replaceTargets.push({
          range,
          prefix: item.prefix,
          ooxml: range.getOoxml()
        });
      }
      await context.sync();

      for (const target of replaceTargets) {
        if (findWordFields(target.ooxml.value).length > 0) continue;
        const replacement = target.range.insertText(target.prefix, Word.InsertLocation.replace);
        replacement.font.bold = true;
      }
      await context.sync();
    });
  }

  /** Convert only parseable Roman table labels, preserving the caption body. */
  static async normalizeTableCaptionPrefixes(indices: number[]): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      const entries: Array<{
        paragraph: Word.Paragraph;
        label: string;
        replacement: string;
        ranges?: Word.RangeCollection;
      }> = [];
      for (const index of [...new Set(indices)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) throw new Error(`Paragraph ${index + 1} is no longer available. Revalidate the document.`);
        paragraph.load('text');
        entries.push({ paragraph, label: '', replacement: '' });
      }
      await context.sync();

      for (const entry of entries) {
        const match = /^(\s*)(TABLE\s+([IVXLCDM]+)\s*\.)/i.exec(entry.paragraph.text || '');
        const number = match ? romanNumeralToNumber(match[3]) : null;
        if (!match || number === null) continue;
        entry.label = match[2];
        entry.replacement = `TABLE ${number}.`;
        entry.ranges = paragraphContentRange(entry.paragraph).search(entry.label, {
          matchCase: true,
          matchWholeWord: false
        });
        entry.ranges.load('items');
      }
      await context.sync();

      const targets: Array<{
        range: Word.Range;
        replacement: string;
        ooxml: OfficeExtension.ClientResult<string>;
      }> = [];
      for (const entry of entries) {
        const range = entry.ranges?.items[0];
        if (!range) continue;
        targets.push({ range, replacement: entry.replacement, ooxml: range.getOoxml() });
      }
      await context.sync();

      for (const target of targets) {
        if (findWordFields(target.ooxml.value).length > 0) continue;
        target.range.insertText(target.replacement, Word.InsertLocation.replace);
      }
      await context.sync();
    });
  }

  /** Bold only the exact leading TABLE n. range; never the caption body. */
  static async applyTableCaptionPrefixBold(indices: number[]): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      const entries: Array<{
        paragraph: Word.Paragraph;
        prefix: string;
        ranges?: Word.RangeCollection;
      }> = [];
      for (const index of [...new Set(indices)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) {
          throw new Error(`Table caption paragraph ${index + 1} is no longer available. Revalidate the document.`);
        }
        paragraph.load('text');
        entries.push({ paragraph, prefix: '' });
      }
      await context.sync();

      for (const entry of entries) {
        const prefix = extractCaptionPrefix(entry.paragraph.text || '', 'table');
        if (!prefix) {
          throw new Error('The exact TABLE n. prefix is no longer available. Revalidate the document.');
        }
        entry.prefix = prefix;
        entry.ranges = paragraphContentRange(entry.paragraph).search(prefix, {
          matchCase: true,
          matchWholeWord: false
        });
        entry.ranges.load('items');
      }
      await context.sync();

      const targets: Array<{ range: Word.Range; prefix: string; ooxml: OfficeExtension.ClientResult<string> }> = [];
      for (const entry of entries) {
        const range = entry.ranges?.items[0];
        if (!range) {
          throw new Error(`Table caption prefix "${entry.prefix}" is no longer available. Revalidate the document.`);
        }
        targets.push({ range, prefix: entry.prefix, ooxml: range.getOoxml() });
      }
      await context.sync();

      for (const target of targets) {
        const fields = findWordFields(target.ooxml.value);
        const hasFieldMarkup = /<(?:[A-Za-z_][\w.-]*:)?(?:fldSimple|fldChar|instrText)\b/i.test(target.ooxml.value);
        const onlyRecognizedTableSequenceFields = fields.length > 0
          && fields.every(field => isTableSequenceFieldInstruction(field.instruction));
        if (hasFieldMarkup && !onlyRecognizedTableSequenceFields) continue;
        target.range.font.bold = true;
      }
      await context.sync();
    });
  }

  /**
   * Apply the complete Reference formatting operation as one transaction.
   * Manual number prefixes are removed only from their exact leading range.
   * When a direct indentation override must be cleared, the paragraph OOXML
   * is round-tripped with only w:ind removed, preserving runs, fields,
   * hyperlinks, and inline emphasis.
   */
  static async applyReferenceFormatting(
    allReferenceIndices: number[],
    options: {
      applyNumbering?: boolean;
      /** Only these Reference layers are allowed to be mutated. */
      properties?: string[];
    } = {}
  ): Promise<void> {
    const orderedReferenceIndices = [...new Set(allReferenceIndices)].sort((a, b) => a - b);
    if (!orderedReferenceIndices.length) return;
    const applyNumbering = options.applyNumbering !== false;
    const requestedProperties = new Set(options.properties || []);
    const allFormattingRequested = options.properties === undefined;
    const styleMetadataRequested = allFormattingRequested || requestedProperties.has('reference.style');
    const fontRequested = allFormattingRequested || requestedProperties.has('reference.style.fontName');
    const sizeRequested = allFormattingRequested || requestedProperties.has('reference.style.fontSize');
    const alignmentRequested = allFormattingRequested || requestedProperties.has('reference.style.alignment');
    const leftIndentRequested = allFormattingRequested || requestedProperties.has('reference.style.leftIndent');
    const hangingIndentRequested = allFormattingRequested || requestedProperties.has('reference.style.hangingIndent');
    const referenceIndentationRequested = leftIndentRequested || hangingIndentRequested;
    const anyParagraphStyleRequested = styleMetadataRequested || fontRequested || sizeRequested
      || alignmentRequested || leftIndentRequested || hangingIndentRequested;
    if (applyNumbering && (typeof Office === 'undefined'
      || !Office.context?.requirements?.isSetSupported('WordApiDesktop', '1.3'))) {
      throw new Error('Reference automatic numbering requires WordApiDesktop 1.3 or later on this host. Review the References section manually.');
    }

    const gapIndices = applyNumbering
      ? orderedReferenceIndices
        .slice(1)
        .flatMap((index, offset) => {
          const previousIndex = orderedReferenceIndices[offset];
          return Array.from({ length: Math.max(0, index - previousIndex - 1) }, (_, gapOffset) =>
            previousIndex + gapOffset + 1
          );
        })
      : [];
    if (gapIndices.length) {
      const layout = await DocxInspector.inspectCurrentDocument();
      const paragraphsByIndex = new Map((layout.paragraphOoxml || []).map(paragraph => [paragraph.index, paragraph]));
      const unsafeGap = gapIndices.some(index => {
        const paragraph = paragraphsByIndex.get(index);
        const blank = paragraph?.blankParagraph;
        const signals = paragraph?.layoutSignals;
        return !paragraph
          || !blank
          || !blank.whitespaceOnly
          || blank.hasManualBreak
          || blank.hasExplicitPageBreak
          || blank.hasLastRenderedPageBreak
          || blank.hasSectionBreak
          || blank.hasDrawing
          || blank.hasField
          || blank.hasEquation
          || blank.hasNumbering
          || blank.hasBookmarkContent
          || paragraph.hasPageBreakBefore === true
          || paragraph.layoutSensitive === true
          || (paragraph.containingTableIndex !== null
            && paragraph.containingTableIndex !== undefined)
          || signals?.hasDrawing === true
          || signals?.hasAnchoredDrawing === true
          || signals?.hasCustomTabs === true
          || signals?.hasNumbering === true;
      });
      if (!layout.ooxmlAvailable || unsafeGap) {
        throw new Error('A structural paragraph occurs between reference entries. Automatic numbering was withheld for manual review.');
      }
    }

    let stage = 'loading reference paragraphs';
    try {
      await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      const missingIndices = orderedReferenceIndices.filter(index => !paragraphs.items[index]);
      if (missingIndices.length) {
        throw new Error(`Reference paragraph mapping is stale for ${missingIndices.length} paragraph(s). Revalidate the document.`);
      }
      let refs = orderedReferenceIndices.map(index => paragraphs.items[index]);
      if (!refs.length) throw new Error('No reference paragraphs were available. Revalidate the document.');
      for (const p of refs) {
        p.load('text,isListItem,style');
      }
      const gapParagraphs = applyNumbering
        ? orderedReferenceIndices
          .slice(1)
          .flatMap((index, offset) => {
            const previousIndex = orderedReferenceIndices[offset];
            return Array.from({ length: Math.max(0, index - previousIndex - 1) }, (_, gapOffset) =>
              paragraphs.items[previousIndex + gapOffset + 1]
            ).filter((paragraph): paragraph is Word.Paragraph => Boolean(paragraph));
          })
        : [];
      for (const p of gapParagraphs) p.load('text');
      const styles = context.document.styles;
      const referenceStyle = styles.getByNameOrNullObject('Reference');
      referenceStyle.load('isNullObject,type,nameLocal,baseStyle,nextParagraphStyle,font/name,font/size,paragraphFormat/alignment,paragraphFormat/leftIndent,paragraphFormat/firstLineIndent');
      const paragraphStyle = styles.getByNameOrNullObject('Paragraph');
      paragraphStyle.load('isNullObject');
      stage = 'loading Reference style and list metadata';
      await context.sync();

      // listString is the displayed value of the existing Word list item. It
      // is separate from the paragraph text and must be captured before any
      // list operation can rebuild the numbering.
      if (applyNumbering) {
        for (const p of refs) {
          if (p.isListItem) p.listItem.load('listString');
        }
        await context.sync();
      }

      if (anyParagraphStyleRequested) {
        if (referenceStyle.isNullObject) {
          throw new Error('The required Reference paragraph style is unavailable. Reference formatting was withheld for manual review.');
        }
        if (paragraphStyle.isNullObject) {
          throw new Error('The required Paragraph base style is unavailable. Reference formatting was withheld for manual review.');
        }
        if (typeof Office === 'undefined'
          || !Office.context?.requirements?.isSetSupported('WordApi', '1.6')) {
          throw new Error('Reference style editing requires WordApi 1.6 or later on this host. Review the References section manually.');
        }
      }

      let hasManualNumbers = false;
      let existingAutomaticListIsCorrect = false;
      if (applyNumbering) {
        const nonBlankGaps = gapParagraphs.filter(p => (p.text || '').trim().length > 0);
        if (nonBlankGaps.length) {
          throw new Error('The References section contains non-reference content between entries. Automatic numbering was withheld for manual review.');
        }
        const manualNumbers = refs.map(p => {
          const match = /^\s*(?:\[\s*(\d+)\s*\]|(\d+)\s*[.)])\s+/.exec(p.text || '');
          return match ? Number(match[1] || match[2]) : null;
        });
        hasManualNumbers = manualNumbers.some(number => number !== null);
        if (hasManualNumbers) {
          if (refs.some(p => p.isListItem)) {
            throw new Error('The References section mixes automatic and manual numbering. Automatic conversion was withheld for manual review.');
          }
          if (manualNumbers.some(number => number === null)) {
            throw new Error('The References section contains an ambiguous mixture of numbered and unnumbered entries. Automatic conversion was withheld for manual review.');
          }
          const nonSequential = manualNumbers.some((number, index) => number !== index + 1);
          if (nonSequential) {
            throw new Error('Manual reference numbers are not sequential from 1. Automatic conversion was withheld to protect in-text reference identities.');
          }
        }

        const hasAutomaticNumbers = refs.some(p => p.isListItem);
        if (hasAutomaticNumbers) {
          const automaticNumbers = refs.map(p => {
            if (!p.isListItem) return null;
            const displayed = (p.listItem.listString || '').trim();
            const match = /^(\d+)\.$/.exec(displayed);
            return match ? Number(match[1]) : null;
          });
          if (automaticNumbers.some(number => number === null)) {
            throw new Error('Existing automatic reference numbering could not be read reliably. Automatic conversion was withheld to protect in-text reference identities.');
          }
          const nonSequentialAutomatic = automaticNumbers.some((number, index) => number !== index + 1);
          if (nonSequentialAutomatic) {
            throw new Error('Existing automatic reference numbers are not sequential from 1. Automatic conversion was withheld to protect in-text reference identities.');
          }
          existingAutomaticListIsCorrect = true;
        }
      }

      // Complete numbering-identity preflight first.  Only after it succeeds
      // do we inspect the coupled style indentation, so an unsafe list is
      // rejected before any style or paragraph XML operation is attempted.
      const normalizedStyleLeft = referenceIndentationRequested
        ? normalizedReferenceStyleLeftIndent(
          referenceStyle.paragraphFormat?.leftIndent,
          referenceStyle.paragraphFormat?.firstLineIndent
        )
        : { value: 0, representation: 'userFacing' as const };
      if (referenceIndentationRequested && normalizedStyleLeft.representation === 'unknown') {
        throw new Error('Reference style indentation representation could not be determined safely. Automatic indentation correction was withheld for manual review.');
      }

      // Direct w:ind properties can override either the Reference style or
      // the list-level positions. Remove only that direct property through
      // the paragraph OOXML so automatic-list indentation remains owned by
      // the list definition and reference typography/content is preserved.
      const paragraphXmlTargets = (leftIndentRequested || hangingIndentRequested)
        ? refs.map((paragraph, position) => ({
        paragraph,
        paragraphIndex: orderedReferenceIndices[position],
        ooxml: paragraph.getOoxml()
        }))
        : [];
      stage = 'checking direct Reference indentation overrides';
      await context.sync();
      const paragraphIndentRepairs = paragraphXmlTargets.flatMap(target => {
        if (typeof target.ooxml.value !== 'string' || !target.ooxml.value.length) {
          throw new Error('Reference paragraph OOXML could not be read reliably. Formatting was withheld for manual review.');
        }
        const repair = repairParagraphIndentOoxml(
          target.ooxml.value,
          target.paragraphIndex,
          target.paragraph.text
        );
        return repair.changed
          ? [{ paragraph: target.paragraph, repaired: repair.xml }]
          : [];
      });
      for (const repair of paragraphIndentRepairs) {
        repair.paragraph.insertOoxml(repair.repaired, Word.InsertLocation.replace);
      }
      if (paragraphIndentRepairs.length) {
        stage = 'applying direct Reference indentation repair';
        await context.sync();
        refs = orderedReferenceIndices.map(index => paragraphs.items[index]);
        for (const p of refs) p.load('text,isListItem,style');
        await context.sync();
      }

      const prefixSearches: Array<{ paragraph: Word.Paragraph; ranges: Word.RangeCollection }> = [];
      if (applyNumbering && hasManualNumbers) {
        for (const p of refs) {
          const match = /^\s*(?:\[\s*\d+\s*\]|\d+\s*[.)])\s+/.exec(p.text || '');
          if (!match) continue;
          const prefix = match[0];
          const ranges = paragraphContentRange(p).search(prefix, {
            matchCase: true,
            matchWholeWord: false
          });
          ranges.load('items');
          prefixSearches.push({ paragraph: p, ranges });
        }
      }
      stage = 'mapping manual Reference number prefixes';
      await context.sync();

      const manualPrefixTargets: Array<{
        range: Word.Range;
        ooxml: OfficeExtension.ClientResult<string>;
      }> = [];
      for (const search of prefixSearches) {
        const prefixRange = search.ranges.items[0];
        if (!prefixRange) {
          throw new Error('A manual reference number prefix could not be mapped to an exact Word range. Automatic correction was withheld for manual review.');
        }
        manualPrefixTargets.push({ range: prefixRange, ooxml: prefixRange.getOoxml() });
      }
      await context.sync();

      if (manualPrefixTargets.some(target => findWordFields(target.ooxml.value).length > 0)) {
        throw new Error('A reference number is field-backed or controlled by citation-management software. Automatic reference formatting was withheld for manual review.');
      }
      if (applyNumbering) {
        const first = refs[0];
        const last = refs[refs.length - 1];
        const referenceRange = paragraphContentRange(first).expandTo(paragraphContentRange(last));
        stage = 'configuring Reference numbering';
        // A correct existing Word list must not be torn down and rebuilt.  It
        // may still be configured when a list-layer rule failed, but preserve
        // its list identity and existing relationships.
        if (!existingAutomaticListIsCorrect) {
          referenceRange.listFormat.applyNumberDefault('Word2000');
        }
        const list = referenceRange.listFormat.list;
        list.setLevelNumbering(0, 'Arabic', [0, '.']);
        list.setLevelStartingNumber(0, 1);
        list.setLevelAlignment(0, 'Left');
        // Text starts at 0.50 in and the number is aligned at 0.25 in.
        list.setLevelIndents(0, 36, -18);
        // The range operation necessarily spans paragraph gaps. Remove any
        // generated list number from blank separators so they remain blank
        // and do not consume a bibliography number.
        for (const blank of gapParagraphs) {
          paragraphContentRange(blank).listFormat.removeNumbers('Paragraph');
        }

        // Synchronize and verify the replacement list before removing any
        // visible manual prefixes. Office.js is not transactional: if this
        // sync or verification fails, the list may be partially applied, but
        // the original identifiers remain available for manual recovery.
        await context.sync();
        for (const p of refs) p.load('text,isListItem');
        stage = 'verifying Reference numbering';
        await context.sync();
        for (const p of refs) {
          if (p.isListItem) p.listItem.load('listString');
        }
        await context.sync();
        const verifiedList = refs.every((p, index) => {
          const displayed = p.isListItem ? (p.listItem.listString || '').trim() : '';
          return /^(\d+)\.$/.test(displayed)
            && Number(displayed.slice(0, -1)) === index + 1;
        });
        if (!verifiedList) {
          throw new Error('Reference automatic numbering was applied but could not be verified as 1., 2., 3. in Word. Manual review is required; visible reference identifiers were preserved.');
        }

        const postListPrefixSearches: Array<{ paragraph: Word.Paragraph; ranges: Word.RangeCollection }> = [];
        if (hasManualNumbers) {
          for (const p of refs) {
            const match = /^\s*(?:\[\s*\d+\s*\]|\d+\s*[.)])\s+/.exec(p.text || '');
            if (!match) continue;
            const ranges = paragraphContentRange(p).search(match[0], {
              matchCase: true,
              matchWholeWord: false
            });
            ranges.load('items');
            postListPrefixSearches.push({ paragraph: p, ranges });
          }
        }
        await context.sync();
        for (const search of postListPrefixSearches) {
          const prefixRange = search.ranges.items[0];
          if (!prefixRange) {
            throw new Error('The verified reference list could not be mapped back to its manual prefixes. Manual review is required; visible reference identifiers were preserved.');
          }
          prefixRange.insertText('', Word.InsertLocation.replace);
        }
        await context.sync();
      }

      stage = 'applying Reference style formatting';
      const contentFormatting = (fontRequested || sizeRequested)
        ? refs.map(p => {
          const range = p.getRange(Word.RangeLocation.content);
          if (typeof (range.font as unknown as { load?: unknown }).load === 'function') {
            range.font.load('name,size');
          }
          return { range, font: range.font };
        })
        : [];
      if (contentFormatting.length) await context.sync();
      if (styleMetadataRequested) {
        if (String(referenceStyle.baseStyle || '').toLowerCase() !== 'paragraph') {
          referenceStyle.baseStyle = 'Paragraph';
        }
        if (String(referenceStyle.nextParagraphStyle || '').toLowerCase() !== 'reference') {
          referenceStyle.nextParagraphStyle = 'Reference';
        }
      }
      if (fontRequested && String(referenceStyle.font.name || '').toLowerCase() !== 'times new roman') {
        referenceStyle.font.name = 'Times New Roman';
      }
      if (sizeRequested && !referenceFormattingApproximately(referenceStyle.font.size, 10)) {
        referenceStyle.font.size = 10;
      }
      if (alignmentRequested && String(referenceStyle.paragraphFormat.alignment || '').toLowerCase() !== 'justified') {
        referenceStyle.paragraphFormat.alignment = Word.Alignment.justified;
      }
      if (referenceIndentationRequested && normalizedStyleLeft.value !== null) {
        // Left and first-line indentation form one coupled pair.  Restore
        // both values together so a malformed pair such as 30 / -10 cannot
        // become 10 / -21.6 and leave a non-zero normalized anchor.
        const targetLeft = normalizedStyleLeft.representation === 'raw'
          ? AIP_RAW_REFERENCE_LEFT_INDENT_PT
          : USER_REFERENCE_LEFT_INDENT_PT;
        const targetFirstLine = normalizedStyleLeft.representation === 'raw'
          ? AIP_RAW_REFERENCE_FIRST_LINE_INDENT_PT
          : USER_REFERENCE_FIRST_LINE_INDENT_PT;
        const leftNeedsRepair = !referenceFormattingApproximately(
          referenceStyle.paragraphFormat.leftIndent,
          targetLeft
        );
        const firstLineNeedsRepair = !referenceFormattingApproximately(
          referenceStyle.paragraphFormat.firstLineIndent,
          targetFirstLine
        );
        if (leftNeedsRepair) {
          referenceStyle.paragraphFormat.leftIndent = targetLeft;
        }
        if (firstLineNeedsRepair) {
          referenceStyle.paragraphFormat.firstLineIndent = targetFirstLine;
        }
      }
      if (anyParagraphStyleRequested) {
        for (const [index, p] of refs.entries()) {
          if (String(p.style || '').toLowerCase() !== 'reference') p.style = 'Reference';
          if (alignmentRequested) p.alignment = Word.Alignment.justified;
          const contentFont = contentFormatting[index]?.font;
          if (fontRequested && String(contentFont?.name || '').toLowerCase() !== 'times new roman') {
            contentFont.name = 'Times New Roman';
          }
          if (sizeRequested && !referenceFormattingApproximately(contentFont?.size, 10)) {
            contentFont.size = 10;
          }
        }
      }
      if (anyParagraphStyleRequested) await context.sync();
      });
    } catch (error) {
      throw formatReferenceFormattingError(stage, error);
    }
  }

  /** Backward-compatible name for callers that only request numbering. */
  static async normalizeReferenceNumbering(allReferenceIndices: number[]): Promise<void> {
    return this.applyReferenceFormatting(allReferenceIndices);
  }

  static async applyParagraphProperty(indices: number[], property: string, expected: string | number | boolean): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      for (const index of indices) {
        const p = paragraphs.items[index];
        if (!p) {
          throw new Error(`Paragraph ${index + 1} is no longer available. Revalidate the document.`);
        }
        const content = paragraphContentRange(p);

        switch (property) {
          case 'font.name':
            content.font.name = String(expected);
            break;
          case 'font.size':
            content.font.size = Number(expected);
            break;
          case 'font.bold':
            content.font.bold = Boolean(expected);
            break;
          case 'font.italic':
            content.font.italic = Boolean(expected);
            break;
          case 'paragraph.alignment':
            p.alignment = this.toWordAlignment(String(expected));
            break;
          case 'paragraph.firstLineIndent':
            p.firstLineIndent = Number(expected);
            break;
          case 'paragraph.leftIndent':
            p.leftIndent = Number(expected);
            break;
          case 'paragraph.rightIndent':
            p.rightIndent = Number(expected);
            break;
          case 'paragraph.spaceBefore':
            p.spaceBefore = Number(expected);
            break;
          case 'paragraph.spaceAfter':
            p.spaceAfter = Number(expected);
            break;
          case 'paragraph.lineSpacing':
            // Office.js stores this value in points while the strategy uses a
            // normalized line multiple (1.0 = single spacing).
            p.lineSpacing = Number(expected) * 12;
            break;
          case 'text.uppercase': {
            const range = paragraphContentRange(p);
            range.load('text');
            await context.sync();
            range.insertText(range.text.toUpperCase(), Word.InsertLocation.replace);
            break;
          }
          default:
            throw new Error(`The correction property '${property}' is not supported.`);
        }
      }

      await context.sync();
    });
  }

  /** Apply both front-matter paragraph spacing values in one Office.js batch. */
  static async applyParagraphSpacing(indices: number[], before: number, after: number): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      for (const index of [...new Set(indices)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) {
          throw new Error(`Paragraph ${index + 1} is no longer available. Revalidate the document.`);
        }
        paragraph.spaceBefore = before;
        paragraph.spaceAfter = after;
      }
      await context.sync();
    });
  }

  /** Change only paragraph space-before, preserving the existing space-after value. */
  static async applyParagraphSpacingBefore(indices: number[], before: number): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      for (const index of [...new Set(indices)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) {
          throw new Error(`Paragraph ${index + 1} is no longer available. Revalidate the document.`);
        }
        paragraph.spaceBefore = before;
      }
      await context.sync();
    });
  }

  /** Bold only the exact leading Abstract. range; never the abstract body. */
  static async applyAbstractPrefixBold(indices: number[]): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      const entries: Array<{
        paragraph: Word.Paragraph;
        prefix: string;
        ranges?: Word.RangeCollection;
      }> = [];
      for (const index of [...new Set(indices)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) {
          throw new Error(`Abstract paragraph ${index + 1} is no longer available. Revalidate the document.`);
        }
        paragraph.load('text');
        entries.push({ paragraph, prefix: '' });
      }
      await context.sync();

      for (const entry of entries) {
        const prefix = extractCaptionPrefix(entry.paragraph.text || '', 'abstract');
        if (!prefix) {
          throw new Error('The exact Abstract. prefix is no longer available. Revalidate the document.');
        }
        entry.prefix = prefix;
        entry.ranges = paragraphContentRange(entry.paragraph).search(prefix, {
          matchCase: true,
          matchWholeWord: false
        });
        entry.ranges.load('items');
      }
      await context.sync();

      const targets: Array<{
        range: Word.Range;
        prefix: string;
        ooxml: OfficeExtension.ClientResult<string>;
      }> = [];
      for (const entry of entries) {
        const range = entry.ranges?.items[0];
        if (!range) {
          throw new Error(`Abstract prefix "${entry.prefix}" is no longer available. Revalidate the document.`);
        }
        targets.push({ range, prefix: entry.prefix, ooxml: range.getOoxml() });
      }
      await context.sync();

      for (const target of targets) {
        if (findWordFields(target.ooxml.value).length > 0) {
          throw new Error('The Abstract. prefix contains field-backed content and was not changed.');
        }
        target.range.font.bold = true;
      }
      await context.sync();
    });
  }

  /** Apply a portable built-in Word paragraph style after a fresh safety check. */
  static async applyBuiltInParagraphStyle(
    indices: number[],
    styleName: Word.BuiltInStyleName = BUILT_IN_HEADING1_STYLE
  ): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      for (const index of [...new Set(indices)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) {
          throw new Error(`Paragraph ${index + 1} is no longer available. Revalidate the document.`);
        }
        paragraph.styleBuiltIn = styleName;
      }

      await context.sync();
    });
  }

  /** Apply only the requested typography properties to paragraph content. */
  static async applyParagraphTypography(indices: number[], options: ParagraphTypographyOptions): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      for (const index of [...new Set(indices)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) {
          throw new Error(`Paragraph ${index + 1} is no longer available. Revalidate the document.`);
        }
        const content = paragraphContentRange(paragraph);
        if (options.fontName !== undefined) content.font.name = options.fontName;
        if (options.fontSize !== undefined) content.font.size = options.fontSize;
        if (options.bold !== undefined) content.font.bold = options.bold;
        if (options.italic !== undefined) content.font.italic = options.italic;
      }

      await context.sync();
    });
  }

  /**
   * Apply author-name size without enlarging raised affiliation/email markers.
   *
   * Word often stores these markers with w:position rather than w:vertAlign
   * and without a direct w:sz. The full paragraph assignment is still useful
   * for the name itself, so record the marker ranges' live sizes first and
   * restore those sizes after applying the requested author-name size.
   */
  static async applyAuthorNameTypographyPreservingMarkers(
    indices: number[],
    options: ParagraphTypographyOptions
  ): Promise<void> {
    const source = await this.snapshot();
    const markerTextsByIndex = new Map<number, Array<{ text: string; fallbackSize: number | null }>>();
    for (const paragraph of source.paragraphs) {
      if (!indices.includes(paragraph.index) || paragraph.classification !== 'authorName') continue;
      const markerTexts = groupRaisedAssociationMarkerRuns(paragraph.effectiveRuns || []);
      if (markerTexts.length) markerTextsByIndex.set(paragraph.index, markerTexts);
    }

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      const entries: Array<{
        paragraph: Word.Paragraph;
        markers: Array<{
          ranges: Word.RangeCollection;
          fallbackSize: number | null;
        }>;
      }> = [];
      for (const index of [...new Set(indices)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) {
          throw new Error(`Paragraph ${index + 1} is no longer available. Revalidate the document.`);
        }
        const markers: Array<{
          ranges: Word.RangeCollection;
          fallbackSize: number | null;
        }> = [];
        for (const marker of markerTextsByIndex.get(index) || []) {
          const ranges = paragraphContentRange(paragraph).search(marker.text, {
            matchCase: true,
            matchWholeWord: false
          });
          ranges.load('items');
          markers.push({ ranges, fallbackSize: marker.fallbackSize });
        }
        entries.push({ paragraph, markers });
      }
      await context.sync();

      const markerSizes: Array<{ range: Word.Range; size: number }> = [];
      for (const entry of entries) {
        for (const marker of entry.markers) {
          for (const range of marker.ranges.items) {
            range.font.load('size');
            markerSizes.push({ range, size: marker.fallbackSize ?? 0 });
          }
        }
      }
      await context.sync();

      for (const marker of markerSizes) {
        if (typeof marker.range.font.size === 'number' && Number.isFinite(marker.range.font.size)) {
          marker.size = marker.range.font.size;
        }
      }

      for (const entry of entries) {
        const content = paragraphContentRange(entry.paragraph);
        if (options.fontName !== undefined) content.font.name = options.fontName;
        if (options.fontSize !== undefined) content.font.size = options.fontSize;
        if (options.bold !== undefined) content.font.bold = options.bold;
        if (options.italic !== undefined) content.font.italic = options.italic;
      }
      await context.sync();

      for (const marker of markerSizes) {
        if (marker.size > 0) marker.range.font.size = marker.size;
      }
      await context.sync();
    });
  }

  /** Set author association text italics without changing a leading superscript marker. */
  static async applyAuthorTextItalic(
    indices: number[],
    target: 'authorAffiliation' | 'authorEmail',
    italic: boolean
  ): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      const entries: Array<{
        paragraph: Word.Paragraph;
        markerRanges?: Word.RangeCollection;
        bodySearches: Word.RangeCollection[];
        fallbackWholeParagraph: boolean;
      }> = [];
      for (const index of [...new Set(indices)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) continue;
        paragraph.load('text');
        entries.push({ paragraph, bodySearches: [], fallbackWholeParagraph: false });
      }
      await context.sync();

      for (const entry of entries) {
        const text = entry.paragraph.text || '';
        if (target === 'authorAffiliation') {
          const marker = text.match(/^\s*(\d+(?:\s*[,;]\s*\d+)*)/)?.[1] || null;
          if (marker) {
            const ranges = paragraphContentRange(entry.paragraph).search(marker, {
              matchCase: true,
              matchWholeWord: false
            });
            ranges.load('items');
            entry.markerRanges = ranges;
          } else {
            entry.fallbackWholeParagraph = true;
          }
          continue;
        }

        const bodyTexts = authorItalicBodyTexts(text, target);
        if (!bodyTexts.length) {
          const normalized = normalizeLogicalLineBreaks(text);
          const hasAssociationMarker = normalized
            .split('\n')
            .some(line => /^\s*[A-Za-z]\)/.test(line));
          // If a marker is missing, there is no association span to protect;
          // only then is it safe to use the whole plain-email paragraph as the
          // fallback. A marker-only line must not cause the marker itself to
          // become italic.
          entry.fallbackWholeParagraph = !hasAssociationMarker
            && /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i.test(normalized);
        }
        for (const bodyText of bodyTexts) {
          const ranges = paragraphContentRange(entry.paragraph).search(bodyText, {
            matchCase: true,
            matchWholeWord: false
          });
          ranges.load('items');
          entry.bodySearches.push(ranges);
        }
      }
      await context.sync();

      for (const entry of entries) {
        if (target === 'authorAffiliation' && entry.markerRanges?.items[0]) {
          const markerRange = entry.markerRanges.items[0];
          const bodyStart = markerRange.getRange('End');
          const paragraphEnd = paragraphContentRange(entry.paragraph).getRange('End');
          bodyStart.expandTo(paragraphEnd).font.italic = italic;
        } else if (entry.bodySearches.length) {
          for (const search of entry.bodySearches) {
            const bodyRange = search.items[0];
            if (bodyRange) bodyRange.font.italic = italic;
          }
        } else if (entry.fallbackWholeParagraph) {
          paragraphContentRange(entry.paragraph).font.italic = italic;
        }
      }
      await context.sync();
    });
  }

  static async applyAuthorAssociationSuperscripts(indices: number[], target: 'authorAffiliation' | 'authorEmail'): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();
      const searches: Word.RangeCollection[] = [];
      for (const index of [...new Set(indices)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) continue;
        paragraph.load('text');
      }
      await context.sync();
      for (const index of [...new Set(indices)]) {
        const paragraph = paragraphs.items[index];
        if (!paragraph) continue;
        for (const marker of associationMarkers(paragraph.text || '', target)) {
          const ranges = paragraphContentRange(paragraph).search(marker, { matchCase: true, matchWholeWord: false });
          ranges.load('items');
          searches.push(ranges);
        }
      }
      await context.sync();
      for (const ranges of searches) {
        for (const marker of ranges.items) marker.font.superscript = true;
      }
      await context.sync();
    });
  }

  /** Replace only whitespace ranges; paragraph runs and their formatting stay intact. */
  static async normalizeParagraphWhitespace(indices: number[], mode: 'leading' | 'internal'): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      const targets = [...new Set(indices)]
        .map(index => paragraphs.items[index])
        .filter((paragraph): paragraph is Word.Paragraph => Boolean(paragraph));
      targets.forEach(paragraph => paragraph.load('text'));
      await context.sync();

      if (mode === 'leading') {
        const searches: Word.RangeCollection[] = [];
        for (const paragraph of targets) {
          const whitespace = (paragraph.text || '').match(/^[ \t]+/)?.[0];
          if (!whitespace) continue;
          const ranges = paragraph.getRange().search(whitespace, { matchCase: true, matchWholeWord: false });
          ranges.load('items');
          searches.push(ranges);
        }
        await context.sync();
        searches.forEach(ranges => {
          const range = ranges.items[0];
          if (range) range.insertText('', Word.InsertLocation.replace);
        });
        await context.sync();
        return;
      }

      // Search/replace in bounded passes so runs are preserved and a sequence
      // of three or more spaces is fully reduced without rewriting the text.
      for (let pass = 0; pass < 16; pass++) {
        const searches: Word.RangeCollection[] = [];
        for (const paragraph of targets) {
          const ranges = paragraph.getRange().search('  ', { matchCase: true, matchWholeWord: false });
          ranges.load('items');
          searches.push(ranges);
        }
        await context.sync();
        let changed = false;
        for (const ranges of searches) {
          for (const range of ranges.items) {
            range.insertText(' ', Word.InsertLocation.replace);
            changed = true;
          }
        }
        if (!changed) break;
        await context.sync();
      }
    });
  }

  /**
   * Apply semantic single spacing only to the affected paragraphs. Word's
   * lineSpacing property is expressed in points, so 12 points is one line for
   * the normal 12-point line grid. The resulting DOCX is verified separately
   * through OOXML; accepting this Office.js assignment alone is not success.
   */
  static async applySingleLineSpacing(indices: number[]): Promise<void> {
    const uniqueIndices = [...new Set(indices)];
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      for (const index of uniqueIndices) {
        const paragraph = paragraphs.items[index];
        if (paragraph) paragraph.lineSpacing = 12;
      }

      await context.sync();
    });
  }

  /** Set one page margin on every section. Word exposes section margins in points. */
  static async applySectionMargin(
    side: 'top' | 'right' | 'bottom' | 'left',
    inches: number
  ): Promise<void> {
    if (!Number.isFinite(inches) || inches < 0) {
      throw new Error('Section margins must be non-negative finite values.');
    }

    await Word.run(async (context) => {
      const sections = context.document.sections;
      sections.load('items');
      await context.sync();

      if (!sections.items.length) {
        throw new Error('No document sections were available for margin correction.');
      }

      const points = inches * 72;
      for (const section of sections.items) {
        switch (side) {
          case 'top': section.pageSetup.topMargin = points; break;
          case 'right': section.pageSetup.rightMargin = points; break;
          case 'bottom': section.pageSetup.bottomMargin = points; break;
          case 'left': section.pageSetup.leftMargin = points; break;
        }
      }
      await context.sync();
    });
  }

  /** Set US Letter paper on every section while preserving portrait/landscape orientation. */
  static async applyLetterPageSize(): Promise<void> {
    if (typeof Office === 'undefined'
      || !Office.context?.requirements?.isSetSupported('WordApiDesktop', '1.3')) {
      throw new Error('US Letter page-size editing requires WordApiDesktop 1.3 or later on this host. Review the page setup manually.');
    }

    await Word.run(async (context) => {
      const sections = context.document.sections;
      sections.load('items');
      await context.sync();
      if (!sections.items.length) throw new Error('No document sections were available for page-size correction.');

      for (const section of sections.items) {
        section.pageSetup.load('orientation,pageWidth,pageHeight');
      }
      await context.sync();

      for (const section of sections.items) {
        const pageSetup = section.pageSetup;
        const landscape = String(pageSetup.orientation).toLowerCase() === 'landscape'
          || pageSetup.pageWidth > pageSetup.pageHeight;
        pageSetup.set({
          paperSize: Word.PaperSize.letter,
          orientation: landscape ? Word.PageOrientation.landscape : Word.PageOrientation.portrait
        });
      }
      await context.sync();
    });
  }

  /** Remove only page-number fields, preserving surrounding header/footer text. */
  static async removePageNumberFields(): Promise<void> {
    if (!Office.context.requirements.isSetSupported('WordApi', '1.5')) {
      throw new Error('Removing page-number fields requires WordApi 1.5 or later on this host.');
    }

    await Word.run(async (context) => {
      const bodies: Word.Body[] = [context.document.body];
      const sections = context.document.sections;
      sections.load('items');
      await context.sync();

      const types: Word.HeaderFooterType[] = [
        Word.HeaderFooterType.primary,
        Word.HeaderFooterType.firstPage,
        Word.HeaderFooterType.evenPages
      ];
      for (const section of sections.items) {
        for (const type of types) {
          bodies.push(section.getHeader(type), section.getFooter(type));
        }
      }

      const fieldCollections = bodies.map(body => {
        const fields = body.fields;
        fields.load('items');
        return fields;
      });
      await context.sync();

      const fields = fieldCollections.flatMap(collection => collection.items);
      for (const field of fields) field.load('type,code');
      await context.sync();

      for (const field of fields) {
        const kind = classifyOfficeField(String(field.type || ''), String(field.code || ''));
        if (kind !== 'OTHER') {
          field.delete();
        }
      }
      await context.sync();
    });
  }

  /** Add the two semantic equation layout tabs without replacing equation text. */
  static async applyEquationTabs(indices: number[]): Promise<void> {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items');
      await context.sync();
      const requests = [...new Set(indices)]
        .map(index => paragraphs.items[index])
        .filter((paragraph): paragraph is Word.Paragraph => Boolean(paragraph))
        .map(paragraph => ({ paragraph, result: paragraph.getRange().getOoxml() }));
      await context.sync();
      for (const request of requests) {
        request.paragraph.getRange().insertOoxml(
          withEquationTabs(request.result.value),
          Word.InsertLocation.replace
        );
      }
      await context.sync();
    });
  }

  /** Add only the requested table rules; content and table layout remain untouched. */
  static async applyTableRule(
    indices: number[],
    rule: 'top' | 'bottom' | 'header' | 'vertical' | 'extraHorizontal',
    exactTableIndices?: number[],
    headerRowCounts?: number[],
    captionRowIndices?: Array<number | null>,
    headerStartRowIndices?: Array<number | null>,
    headerEndRowIndices?: Array<number | null>
  ): Promise<void> {
    await Word.run(async (context) => {
      type TableFixTarget = {
        tableIndex: number;
        headerRowCount: number;
        captionRowIndex: number | null;
        headerStartRowIndex: number | null;
        headerEndRowIndex: number | null;
      };
      const paragraphs = context.document.body.paragraphs;
      const tables = context.document.body.tables;
      paragraphs.load('items');
      tables.load('items');
      await context.sync();
      paragraphs.items.forEach(paragraph => paragraph.load('text'));
      await context.sync();

      const captionOrdinals = paragraphs.items
        .map((paragraph, index) => ({ paragraph, index }))
        .filter(item => /^\s*TABLE\s+(?:\d+[A-Z]?|[IVXLCDM]+)\s*[.:)-]/i.test(item.paragraph.text || ''))
        .map(item => item.index);
      const rawTargetRecords: TableFixTarget[] = exactTableIndices?.length
        ? exactTableIndices.map((tableIndex, position) => ({
          tableIndex,
          headerRowCount: Math.max(1, headerRowCounts?.[position] || 1),
          captionRowIndex: captionRowIndices?.[position] ?? null,
          headerStartRowIndex: headerStartRowIndices?.[position] ?? null,
          headerEndRowIndex: headerEndRowIndices?.[position] ?? null
        }))
        : [...new Set(indices)]
          .map((index, position) => {
            const tableIndex = captionOrdinals.indexOf(index);
            return tableIndex < 0 ? null : {
              tableIndex,
              headerRowCount: Math.max(1, headerRowCounts?.[position] || 1),
              captionRowIndex: captionRowIndices?.[position] ?? null,
              headerStartRowIndex: headerStartRowIndices?.[position] ?? null,
              headerEndRowIndex: headerEndRowIndices?.[position] ?? null
            };
          })
          .filter((record): record is TableFixTarget => record !== null);
      const uniqueTargetRecords = [...new Map(rawTargetRecords.map(record => [record.tableIndex, record])).values()];
      const targetEntries = uniqueTargetRecords
        .map(record => ({ record, table: tables.items[record.tableIndex] }))
        .filter((entry): entry is { record: TableFixTarget; table: Word.Table } => Boolean(entry.table));
      const targetRecords = targetEntries.map(entry => entry.record);
      const targets = targetEntries.map(entry => entry.table);

      const borderRequests: Word.TableBorder[] = [];
      const restoreBorderRequests: Word.TableBorder[] = [];
      const insideHorizontalBorders: Array<{ table: Word.Table; border: Word.TableBorder; wasPresent: boolean }> = [];
      for (const [tablePosition, table] of targets.entries()) {
        const targetRecord = targetRecords[tablePosition];
        const embeddedTop = rule === 'top'
          && targetRecord?.captionRowIndex !== null
          && targetRecord?.captionRowIndex !== undefined
          && targetRecord?.headerStartRowIndex !== null
          && targetRecord?.headerStartRowIndex !== undefined;
        if (rule === 'header' || rule === 'vertical' || rule === 'extraHorizontal' || embeddedTop) {
          table.rows.load('items');
        } else {
          const border = table.getBorder(rule === 'top' ? Word.BorderLocation.top : Word.BorderLocation.bottom);
          border.load('type');
          borderRequests.push(border);
        }
        if (rule === 'extraHorizontal') {
          const insideHorizontal = table.getBorder(Word.BorderLocation.insideHorizontal);
          insideHorizontal.load('type');
          insideHorizontalBorders.push({ table, border: insideHorizontal, wasPresent: false });
          borderRequests.push(insideHorizontal);
        }
      }
      await context.sync();
      for (const entry of insideHorizontalBorders) {
        entry.wasPresent = entry.border.type !== Word.BorderType.none;
      }

      if (rule === 'header' || rule === 'top') {
        const boundaryTables = rule === 'header' || targets.some((_, index) =>
          targetRecords[index]?.captionRowIndex !== null && targetRecords[index]?.captionRowIndex !== undefined
          && targetRecords[index]?.headerStartRowIndex !== null && targetRecords[index]?.headerStartRowIndex !== undefined
        );
        if (boundaryTables) {
          for (const table of targets) {
            for (const row of table.rows.items) row.cells.load('items');
          }
          await context.sync();
        }
      }

      if (rule === 'header') {
        for (const [tablePosition, table] of targets.entries()) {
          const targetRecord = targetRecords[tablePosition];
          const headerCount = targetRecord?.headerRowCount || 1;
          const headerEnd = targetRecord?.headerEndRowIndex ?? headerCount - 1;
          const row = table.rows.items[headerEnd];
          if (!row) continue;
          const nextRow = table.rows.items[headerEnd + 1];
          const borders = [row.getBorder(Word.BorderLocation.bottom)];
          if (nextRow) borders.push(nextRow.getBorder(Word.BorderLocation.top));
          for (const border of borders) {
            border.load('type');
            borderRequests.push(border);
          }
          for (const cell of row.cells.items) {
            const border = cell.getBorder(Word.BorderLocation.bottom);
            border.load('type');
            borderRequests.push(border);
          }
          if (nextRow) {
            for (const cell of nextRow.cells.items) {
              const border = cell.getBorder(Word.BorderLocation.top);
              border.load('type');
              borderRequests.push(border);
            }
          }
        }
        await context.sync();
      }
      if (rule === 'top') {
        for (const [tablePosition, table] of targets.entries()) {
          const targetRecord = targetRecords[tablePosition];
          const captionRow = targetRecord?.captionRowIndex;
          const headerStart = targetRecord?.headerStartRowIndex;
          if (captionRow === null || captionRow === undefined || headerStart === null || headerStart === undefined) continue;
          const upperRow = table.rows.items[captionRow];
          const lowerRow = table.rows.items[headerStart];
          if (!upperRow || !lowerRow) continue;
          const upperBorder = upperRow.getBorder(Word.BorderLocation.bottom);
          const lowerBorder = lowerRow.getBorder(Word.BorderLocation.top);
          upperBorder.load('type');
          lowerBorder.load('type');
          borderRequests.push(upperBorder, lowerBorder);
          for (const cell of upperRow.cells.items) {
            const border = cell.getBorder(Word.BorderLocation.bottom);
            border.load('type');
            borderRequests.push(border);
          }
          for (const cell of lowerRow.cells.items) {
            const border = cell.getBorder(Word.BorderLocation.top);
            border.load('type');
            borderRequests.push(border);
          }
        }
        await context.sync();
      }
      if (rule === 'vertical' || rule === 'extraHorizontal') {
        for (const table of targets) {
          for (const row of table.rows.items) row.cells.load('items');
        }
        await context.sync();

        for (const table of targets) {
          const rows = table.rows.items;
          if (rule === 'vertical') {
            const tableBorder = table.getBorder(Word.BorderLocation.insideVertical);
            tableBorder.load('type');
            borderRequests.push(tableBorder);
            const left = table.getBorder(Word.BorderLocation.left);
            const right = table.getBorder(Word.BorderLocation.right);
            left.load('type');
            right.load('type');
            borderRequests.push(left, right);
            for (const row of rows) {
              for (const cell of row.cells.items) {
                const left = cell.getBorder(Word.BorderLocation.left);
                const right = cell.getBorder(Word.BorderLocation.right);
                left.load('type');
                right.load('type');
                borderRequests.push(left, right);
              }
            }
          } else {
            // Keep the required header separator and final outer bottom rule.
            const tablePosition = targets.indexOf(table);
            const targetRecord = targetRecords[tablePosition];
            const headerCount = targetRecord?.headerRowCount || 1;
            const headerStart = targetRecord?.headerStartRowIndex ?? 0;
            const headerEnd = targetRecord?.headerEndRowIndex ?? headerCount - 1;

            const queueBoundaryRestore = (upperRow: Word.TableRow, lowerRow: Word.TableRow) => {
              const upperBorder = upperRow.getBorder(Word.BorderLocation.bottom);
              const lowerBorder = lowerRow.getBorder(Word.BorderLocation.top);
              restoreBorderRequests.push(upperBorder, lowerBorder);
              for (const cell of upperRow.cells.items) {
                restoreBorderRequests.push(cell.getBorder(Word.BorderLocation.bottom));
              }
              for (const cell of lowerRow.cells.items) {
                restoreBorderRequests.push(cell.getBorder(Word.BorderLocation.top));
              }
            };

            // Remove forbidden boundaries inside a multi-row header as well
            // as inside the body. The header/body boundary itself remains.
            for (let lowerIndex = headerStart + 1; lowerIndex <= headerEnd; lowerIndex += 1) {
              const upperRow = rows[lowerIndex - 1];
              const lowerRow = rows[lowerIndex];
              if (!upperRow || !lowerRow) continue;
              const rowBottom = upperRow.getBorder(Word.BorderLocation.bottom);
              const lowerRowTop = lowerRow.getBorder(Word.BorderLocation.top);
              rowBottom.load('type');
              lowerRowTop.load('type');
              borderRequests.push(rowBottom, lowerRowTop);
              for (const cell of upperRow.cells.items) {
                const bottom = cell.getBorder(Word.BorderLocation.bottom);
                bottom.load('type');
                borderRequests.push(bottom);
              }
              for (const cell of lowerRow.cells.items) {
                const top = cell.getBorder(Word.BorderLocation.top);
                top.load('type');
                borderRequests.push(top);
              }
            }

            for (let lowerIndex = headerEnd + 2; lowerIndex < rows.length; lowerIndex++) {
              const upperRow = rows[lowerIndex - 1];
              const lowerRow = rows[lowerIndex];
              const rowBottom = upperRow.getBorder(Word.BorderLocation.bottom);
              const lowerRowTop = lowerRow.getBorder(Word.BorderLocation.top);
              rowBottom.load('type');
              lowerRowTop.load('type');
              borderRequests.push(rowBottom, lowerRowTop);
              const upperCells = upperRow.cells.items;
              const lowerCells = lowerRow.cells.items;
              for (const cell of upperCells) {
                const bottom = cell.getBorder(Word.BorderLocation.bottom);
                bottom.load('type');
                borderRequests.push(bottom);
              }
              for (const cell of lowerCells) {
                const top = cell.getBorder(Word.BorderLocation.top);
                top.load('type');
                borderRequests.push(top);
              }
            }

            // Removing table-level insideH also removes any required rule it
            // supplied. Restore the embedded caption/header and header/body
            // boundaries explicitly after the forbidden rules are cleared.
            const captionRow = targetRecord?.captionRowIndex;
            const insideHWasPresent = insideHorizontalBorders.find(entry => entry.table === table)?.wasPresent === true;
            const restoreBoundaries = tableBoundariesToRestore(
              insideHWasPresent,
              captionRow !== null && captionRow !== undefined
                && headerStart >= 0 && headerStart < rows.length
                && Boolean(rows[captionRow]) && Boolean(rows[headerStart]),
              headerEnd >= 0 && headerEnd + 1 < rows.length
            );
            if (restoreBoundaries.captionHeader && captionRow !== null && captionRow !== undefined && headerStart !== null && headerStart !== undefined) {
                queueBoundaryRestore(rows[captionRow], rows[headerStart]);
            }
            if (restoreBoundaries.headerBody) {
              const firstBodyIndex = headerEnd + 1;
              queueBoundaryRestore(rows[headerEnd], rows[firstBodyIndex]);
            }
          }
        }
        await context.sync();
      }
      const removeRule = rule === 'vertical' || rule === 'extraHorizontal';
      for (const border of borderRequests) {
        // Setting none on the effective border object creates an explicit
        // direct override when the visible border came from a table style.
        // This prevents an inherited TableGrid insideV/insideH rule from
        // reappearing after the correction is revalidated.
        border.type = removeRule
          ? Word.BorderType.none
          : Word.BorderType.single;
      }
      if (rule === 'extraHorizontal') {
        for (const border of restoreBorderRequests) border.type = Word.BorderType.single;
      }
      await context.sync();
    });
  }

  private static toWordAlignment(value: string): Word.Alignment {
    const normalized = value.toLowerCase();
    if (normalized === 'centered' || normalized === 'center') return Word.Alignment.centered;
    if (normalized === 'right') return Word.Alignment.right;
    if (normalized === 'justified' || normalized === 'both') return Word.Alignment.justified;
    return Word.Alignment.left;
  }
}

function withEquationTabs(ooxml: string): string {
  const required = '<w:tab w:val="center" w:pos="4320"/><w:tab w:val="right" w:pos="9242"/>';
  const existing = ooxml.match(/<(?:[A-Za-z_][\w.-]*:)?tabs\b[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?tabs\s*>/i);
  if (existing) {
    const preserved = existing[0]
      .replace(/^<[A-Za-z_][\w.-]*:tabs\b[^>]*>/i, '')
      .replace(/<\/(?:[A-Za-z_][\w.-]*:)?tabs\s*>$/i, '')
      .replace(/<(?:[A-Za-z_][\w.-]*:)?tab\b[^>]*\/\s*>/gi, tag =>
        /(?:w:|\s)val\s*=\s*["'](?:center|right)["']/i.test(tag) ? '' : tag
      );
    return ooxml.replace(existing[0], `<w:tabs>${preserved}${required}</w:tabs>`);
  }
  if (/<(?:[A-Za-z_][\w.-]*:)?pPr\b[^>]*>/i.test(ooxml)) {
    return ooxml.replace(/(<(?:[A-Za-z_][\w.-]*:)?pPr\b[^>]*>)/i, `$1<w:tabs>${required}</w:tabs>`);
  }
  return ooxml.replace(/(<(?:[A-Za-z_][\w.-]*:)?p\b[^>]*>)/i, `$1<w:pPr><w:tabs>${required}</w:tabs></w:pPr>`);
}
