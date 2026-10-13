import {
  DocumentSnapshot,
  FrontMatterTransitionFixTarget,
  FrontMatterType,
  Finding,
  ParagraphSnapshot,
  ReviewKind,
  RuleDefinition,
  StrategyDefinition,
  TableCaptionTableGapAnalysis,
  TableCaptionTableGapIssue,
  ValidationSummary
} from './models';
import { approxEqual, displayValue, isAllCaps, looksMajorTitleCase, normalizeAlignment, paragraphTextFingerprint } from './utils';
import { ApplicabilityAnalyzer } from './ApplicabilityAnalyzer';
import { parseCaptionNumber, romanNumeralToNumber } from './CaptionNumberParser';
import { CanonicalSection, sectionFamilyLabel } from './SectionOntology';

const severityWeight = { critical: 5, major: 3, minor: 1, info: 0.5 } as const;
const SMALL_HEADING_WORDS = new Set([
  'a', 'an', 'the', 'and', 'but', 'or', 'nor', 'for', 'from', 'with', 'in', 'on',
  'at', 'by', 'to', 'of', 'off', 'as', 'via', 'vs', 'versus', 'per'
]);

function isBlankParagraph(paragraph: ParagraphSnapshot): boolean {
  return paragraph.isBlank === true
    || (paragraph.isBlank === undefined && paragraph.text.replace(/\u00A0/g, ' ').trim().length === 0);
}

type ReferenceStyleState = 'reference' | 'wrong' | 'unknown';

function referenceStyleState(paragraph: ParagraphSnapshot): ReferenceStyleState {
  const label = (paragraph.style || paragraph.styleBuiltIn || '').trim();
  if (!label) return 'unknown';
  return /^reference$/i.test(label) ? 'reference' : 'wrong';
}

/** Paragraphs safe for generic prose/layout analysis. */
function isReliableLayoutParagraph(paragraph: ParagraphSnapshot): boolean {
  return paragraph.ooxmlAligned !== false && paragraph.layoutSensitive !== true;
}

export function hasHeading1ParagraphStyle(paragraph: ParagraphSnapshot): boolean {
  return /(?:^|\s)heading\s*1(?:\s|$)/i.test(
    `${paragraph.style} ${paragraph.styleBuiltIn}`
  );
}

/** Return only the first true Heading 1 that follows the document Abstract. */
export function findFirstHeading1AfterAbstract(
  paragraphs: ParagraphSnapshot[]
): ParagraphSnapshot | null {
  const ordered = [...paragraphs].sort((a, b) => a.index - b.index);
  const abstract = ordered.find(paragraph => paragraph.classification === 'abstract');
  if (!abstract) return null;
  return ordered.find(paragraph =>
    paragraph.classification === 'heading1'
    && paragraph.index > abstract.index
  ) || null;
}

const NON_HEADING1_CLASSIFICATIONS = new Set<ParagraphSnapshot['classification']>([
  'title',
  'authorName',
  'authorAffiliation',
  'authorEmail',
  'abstract',
  'figureCaption',
  'tableCaption',
  'equation',
  'reference'
]);

const REQUIRED_SECTION_GROUPS: Record<string, {
  label: string;
  families: readonly CanonicalSection[];
}> = {
  'structure.introductionPresence': {
    label: 'Introduction',
    families: ['introduction']
  },
  'structure.methodSectionPresence': {
    label: 'Methods / approach',
    families: [
      'methodology',
      'materials_methods',
      'experimental_method',
      'proposed_solution',
      'system_architecture',
      'implementation'
    ]
  },
  'structure.resultsSectionPresence': {
    label: 'Results / evaluation',
    families: [
      'results',
      'results_discussion',
      'evaluation',
      'analysis'
    ]
  },
  'structure.conclusionPresence': {
    label: 'Conclusion',
    families: ['conclusion', 'conclusion_future_work']
  }
};

function isRecognizedFirstLevelSection(paragraph: ParagraphSnapshot): boolean {
  if (paragraph.sectionConfidence !== 'high' && paragraph.sectionConfidence !== 'medium') {
    return false;
  }
  if (paragraph.classification === 'heading2' || paragraph.classification === 'heading3') {
    return false;
  }
  return paragraph.classification === 'heading1'
    || (paragraph.sectionConfidence === 'high' && (paragraph.sectionFamilies?.length || 0) > 0);
}

/**
 * Candidate paragraphs for the dedicated Heading 1 style correction. The
 * ontology may identify a known section even when Word assigned Normal; an
 * explicit Heading 2/3 classification always wins and is never promoted.
 */
export function isSafeHeading1StyleCandidate(paragraph: ParagraphSnapshot): boolean {
  if (NON_HEADING1_CLASSIFICATIONS.has(paragraph.classification)) return false;
  if (paragraph.classification === 'heading2' || paragraph.classification === 'heading3') return false;

  const highConfidenceCandidate = paragraph.heading1Confidence === 'high';
  return paragraph.tableNestingLevel === 0
    && paragraph.layoutSensitive !== true
    && paragraph.ooxmlAligned !== false
    && (paragraph.classification === 'heading1' || highConfidenceCandidate);
}

function isRaisedAssociationRun(
  run: NonNullable<ParagraphSnapshot['effectiveRuns']>[number]
): boolean {
  return run.verticalAlign === 'superscript'
    || (run.baselinePositionPt ?? 0) > 0;
}

function figureCaptionLabelSpan(text: string): { start: number; end: number } | null {
  const match = /^(\s*)((?:FIGURE|FIG\.?)\s+\d+[A-Za-z]?(?:\s*[.\-:)])?)/i.exec(text);
  if (!match?.[2]) return null;
  return {
    start: match[1].length,
    end: match[1].length + match[2].length
  };
}

function hasFieldBackedFigureCaptionLabel(paragraph: ParagraphSnapshot): boolean {
  const label = figureCaptionLabelSpan(paragraph.text);
  if (!label || !paragraph.fieldTextSpans?.length) return false;
  return paragraph.fieldTextSpans.some(field => field.start < label.end && field.end > label.start);
}

function hasUnparseableFigureCaptionLabel(paragraph: ParagraphSnapshot): boolean {
  return /^(?:FIGURE|FIG\.?)\b/i.test(paragraph.text.trim())
    && figureCaptionLabelSpan(paragraph.text) === null;
}

function romanTableCaptionLabelSpan(text: string): { start: number; end: number } | null {
  const match = /^(\s*)(TABLE\s+[IVXLCDM]+\s*\.)/i.exec(text);
  if (!match?.[2]) return null;
  return {
    start: match[1].length,
    end: match[1].length + match[2].length
  };
}

function hasFieldBackedRomanTableCaptionLabel(paragraph: ParagraphSnapshot): boolean {
  const label = romanTableCaptionLabelSpan(paragraph.text);
  if (!label || !paragraph.fieldTextSpans?.length) return false;
  return paragraph.fieldTextSpans.some(field => field.start < label.end && field.end > label.start);
}

function tableCaptionPrefixSpan(text: string): { start: number; end: number } | null {
  const match = /^(\s*)(TABLE\s+(?:\d+[A-Za-z]?|[IVXLCDM]+)\s*\.)/i.exec(text);
  if (!match?.[2]) return null;
  return {
    start: match[1].length,
    end: match[1].length + match[2].length
  };
}

function hasFieldBackedTableCaptionPrefix(paragraph: ParagraphSnapshot): boolean {
  return tableCaptionPrefixFieldKind(paragraph) !== 'none';
}

function tableCaptionPrefixFieldKind(paragraph: ParagraphSnapshot): 'none' | 'tableSequence' | 'other' {
  const prefix = tableCaptionPrefixSpan(paragraph.text);
  if (!prefix) return 'other';
  const overlapping = (paragraph.fieldTextSpans || [])
    .filter(field => field.start < prefix.end && field.end > prefix.start);
  if (!overlapping.length) {
    return paragraph.hasField === true ? 'other' : 'none';
  }
  return overlapping.every(field => field.kind === 'tableSequence')
    ? 'tableSequence'
    : 'other';
}

function isSafeTablePrefixBoldCandidate(
  paragraph: ParagraphSnapshot,
  snapshot: DocumentSnapshot
): boolean {
  return tableCaptionPrefixSpan(paragraph.text) !== null
    && tableCaptionPrefixFieldKind(paragraph) !== 'other'
    && paragraph.tableNestingLevel === 0
    && paragraph.layoutSensitive !== true
    && paragraph.ooxmlAligned !== false
    && snapshot.layout.paragraphMappingAligned !== false;
}

function abstractPrefixSpan(text: string): { start: number; end: number } | null {
  const match = /^(\s*)(Abstract\s*\.)/i.exec(text);
  if (!match?.[2]) return null;
  return {
    start: match[1].length,
    end: match[1].length + match[2].length
  };
}

export function hasFieldBackedAbstractPrefix(paragraph: ParagraphSnapshot): boolean {
  const prefix = abstractPrefixSpan(paragraph.text);
  if (!prefix) return false;
  if (paragraph.hasField === true && !paragraph.fieldTextSpans?.length) return true;
  return paragraph.fieldTextSpans?.some(field => field.start < prefix.end && field.end > prefix.start) === true;
}

export function isSafeAbstractPrefixBoldCandidate(
  paragraph: ParagraphSnapshot,
  paragraphMappingReliable = true
): boolean {
  return paragraph.classification === 'abstract'
    && abstractPrefixSpan(paragraph.text) !== null
    && paragraphMappingReliable
    && paragraph.ooxmlAligned !== false
    && paragraph.layoutSensitive !== true
    && !hasFieldBackedAbstractPrefix(paragraph);
}

function isSafeRomanTableCaptionFix(paragraph: ParagraphSnapshot): boolean {
  const match = /^\s*TABLE\s+([IVXLCDM]+)\s*\./i.exec(paragraph.text);
  return match !== null
    && romanNumeralToNumber(match[1]) !== null
    && !hasFieldBackedRomanTableCaptionLabel(paragraph);
}

/**
 * Return only safely removable blank paragraph nodes that sit between two
 * top-level body paragraphs. This intentionally does not infer blank lines
 * from spacing, line breaks, or rendered layout.
 */
export function findExtraBlankBodyParagraphIndices(paragraphs: ParagraphSnapshot[]): number[] {
  const previousMeaningful = (index: number): ParagraphSnapshot | null => {
    for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
      if (!isBlankParagraph(paragraphs[cursor])) return paragraphs[cursor];
    }
    return null;
  };
  const nextMeaningful = (index: number): ParagraphSnapshot | null => {
    for (let cursor = index + 1; cursor < paragraphs.length; cursor += 1) {
      if (!isBlankParagraph(paragraphs[cursor])) return paragraphs[cursor];
    }
    return null;
  };

  const blankRunIsSafe = (index: number): boolean => {
    let start = index;
    let end = index;
    while (start > 0 && isBlankParagraph(paragraphs[start - 1])) start -= 1;
    while (end + 1 < paragraphs.length && isBlankParagraph(paragraphs[end + 1])) end += 1;
    for (let cursor = start; cursor <= end; cursor += 1) {
      const paragraph = paragraphs[cursor];
      if (paragraph.tableNestingLevel !== 0
        || paragraph.isListItem
        || paragraph.blankParagraph?.safeToDelete !== true) {
        return false;
      }
    }
    return true;
  };

  return paragraphs
    .filter((paragraph, index) => {
      if (!isBlankParagraph(paragraph)) return false;
      if (paragraph.tableNestingLevel !== 0 || paragraph.isListItem) return false;
      if (paragraph.blankParagraph?.safeToDelete !== true) return false;
      if (!blankRunIsSafe(index)) return false;

      const previous = previousMeaningful(index);
      const next = nextMeaningful(index);
      return previous?.classification === 'body'
        && next?.classification === 'body'
        && !previous.isListItem
        && !next.isListItem
        && previous.tableNestingLevel === 0
        && next.tableNestingLevel === 0;
    })
    .map(paragraph => paragraph.index);
}

export interface FigureCaptionBodyGapIssue {
  captionParagraphIndex: number;
  bodyParagraphIndex: number;
  blankParagraphIndices: number[];
  blockedByStructuralParagraph?: boolean;
}

export interface FigureCaptionBodyGapAnalysis {
  issues: FigureCaptionBodyGapIssue[];
  applicableCaptionCount: number;
  ignoredCaptionIndices: number[];
  blockedCaptionIndices: number[];
}

export interface FigureSpaceAboveIssue {
  captionParagraphIndex: number;
  figureParagraphIndex: number;
  previousParagraphIndex: number;
  blankParagraphIndices: number[];
  blockedByStructuralParagraph?: boolean;
}

export interface FigureSpaceAboveAnalysis {
  issues: FigureSpaceAboveIssue[];
  applicableFigureCount: number;
  ignoredCaptionIndices: number[];
  blockedCaptionIndices: number[];
}

function isSafeTopLevelBlank(paragraph: ParagraphSnapshot): boolean {
  return isBlankParagraph(paragraph)
    && paragraph.tableNestingLevel === 0
    && !paragraph.isListItem
    && paragraph.blankParagraph?.safeToDelete === true;
}

export type AffiliationEmailBlankLineState =
  | 'not_applicable'
  | 'passed'
  | 'missing'
  | 'extra'
  | 'blocked';

export interface AffiliationEmailBlankLineAnalysis {
  state: AffiliationEmailBlankLineState;
  lastAffiliationParagraphIndex?: number;
  firstEmailParagraphIndex?: number;
  blankParagraphIndices: number[];
  removableBlankParagraphIndices: number[];
  blockedParagraphIndices: number[];
  separatorKind?: 'blankParagraph' | 'leadingManualBreak' | 'none' | 'mixed' | 'unknown';
  leadingManualBreakCount?: number | null;
  reason?: string;
}

export function isStructurallySafeFrontMatterBlank(paragraph: ParagraphSnapshot): boolean {
  const metadata = paragraph.blankParagraph;
  return isSafeTopLevelBlank(paragraph)
    && paragraph.layoutSensitive !== true
    && paragraph.hasPageBreakBefore !== true
    && paragraph.ooxmlAligned !== false
    && metadata?.hasManualBreak === false
    && metadata.hasExplicitPageBreak === false
    && metadata.hasLastRenderedPageBreak === false
    && metadata.hasSectionBreak === false
    && metadata.hasDrawing === false
    && metadata.hasField === false
    && metadata.hasEquation === false
    && metadata.hasNumbering === false
    && metadata.hasBookmarkContent === false
    && paragraph.hasField !== true
    && paragraph.layoutSignals?.hasDrawing !== true
    && paragraph.layoutSignals?.hasAnchoredDrawing !== true
    && paragraph.layoutSignals?.hasNumbering !== true;
}

function isSafeAffiliationEmailBlank(paragraph: ParagraphSnapshot): boolean {
  return isStructurallySafeFrontMatterBlank(paragraph);
}

/**
 * Analyze the custom front-matter separator policy. Only the region between
 * the last affiliation and the first following email is considered. Any
 * intervening content or structural blank makes the relationship unsafe for
 * automatic editing.
 */
export function analyzeAffiliationEmailBlankLine(
  paragraphs: ParagraphSnapshot[]
): AffiliationEmailBlankLineAnalysis {
  const ordered = [...paragraphs].sort((a, b) => a.index - b.index);
  const lastAffiliationPosition = ordered.reduce(
    (last, paragraph, position) => paragraph.classification === 'authorAffiliation' ? position : last,
    -1
  );
  if (lastAffiliationPosition < 0) {
    return {
      state: 'not_applicable',
      blankParagraphIndices: [],
      removableBlankParagraphIndices: [],
      blockedParagraphIndices: [],
      separatorKind: 'none',
      reason: 'No author-affiliation paragraph was detected.'
    };
  }

  const firstEmailPosition = ordered.findIndex(
    (paragraph, position) => position > lastAffiliationPosition && paragraph.classification === 'authorEmail'
  );
  if (firstEmailPosition < 0) {
    return {
      state: 'not_applicable',
      lastAffiliationParagraphIndex: ordered[lastAffiliationPosition].index,
      blankParagraphIndices: [],
      removableBlankParagraphIndices: [],
      blockedParagraphIndices: [],
      separatorKind: 'none',
      reason: 'No author-email paragraph follows the affiliations.'
    };
  }

  const lastAffiliationParagraph = ordered[lastAffiliationPosition];
  const firstEmailParagraph = ordered[firstEmailPosition];
  const between = ordered.slice(lastAffiliationPosition + 1, firstEmailPosition);
  const nonBlank = between.filter(paragraph => !isBlankParagraph(paragraph));
  const contextIndices = [lastAffiliationParagraph.index, firstEmailParagraph.index];
  if (nonBlank.length) {
    return {
      state: 'blocked',
      lastAffiliationParagraphIndex: lastAffiliationParagraph.index,
      firstEmailParagraphIndex: firstEmailParagraph.index,
      blankParagraphIndices: [],
      removableBlankParagraphIndices: [],
      blockedParagraphIndices: nonBlank.map(paragraph => paragraph.index),
      separatorKind: 'unknown',
      reason: 'Non-blank content intervenes between the last affiliation and first email.'
    };
  }

  const blocked = between.filter(paragraph => !isSafeAffiliationEmailBlank(paragraph));
  if (blocked.length) {
    return {
      state: 'blocked',
      lastAffiliationParagraphIndex: lastAffiliationParagraph.index,
      firstEmailParagraphIndex: firstEmailParagraph.index,
      blankParagraphIndices: between.map(paragraph => paragraph.index),
      removableBlankParagraphIndices: [],
      blockedParagraphIndices: blocked.map(paragraph => paragraph.index),
      separatorKind: 'unknown',
      reason: 'One or more empty separator paragraphs contain structural or positioned-layout content.'
    };
  }

  const blankParagraphIndices = between.map(paragraph => paragraph.index);
  // Synthetic/offline snapshots created before this metadata existed are
  // treated as having no leading break. Live snapshots explicitly use null
  // when OOXML could not be aligned, which must remain manual review.
  const leadingManualBreakCount = firstEmailParagraph.leadingManualLineBreakCount === undefined
    ? (firstEmailParagraph.ooxmlAligned === false ? null : 0)
    : firstEmailParagraph.leadingManualLineBreakCount;

  if (leadingManualBreakCount === null) {
    return {
      state: 'blocked',
      lastAffiliationParagraphIndex: lastAffiliationParagraph.index,
      firstEmailParagraphIndex: firstEmailParagraph.index,
      blankParagraphIndices,
      removableBlankParagraphIndices: [],
      blockedParagraphIndices: [firstEmailParagraph.index],
      separatorKind: 'unknown',
      leadingManualBreakCount,
      reason: 'The first author-email paragraph could not be inspected reliably for a leading manual line break.'
    };
  }

  if (leadingManualBreakCount > 1) {
    return {
      state: 'blocked',
      lastAffiliationParagraphIndex: lastAffiliationParagraph.index,
      firstEmailParagraphIndex: firstEmailParagraph.index,
      blankParagraphIndices,
      removableBlankParagraphIndices: [],
      blockedParagraphIndices: [firstEmailParagraph.index],
      separatorKind: 'unknown',
      leadingManualBreakCount,
      reason: `The first author-email paragraph contains ${leadingManualBreakCount} leading manual line breaks; automatic normalization was withheld.`
    };
  }

  const hasLeadingManualBreak = leadingManualBreakCount === 1;
  if (blankParagraphIndices.length === 0 && hasLeadingManualBreak) {
    return {
      state: 'passed',
      lastAffiliationParagraphIndex: lastAffiliationParagraph.index,
      firstEmailParagraphIndex: firstEmailParagraph.index,
      blankParagraphIndices,
      removableBlankParagraphIndices: [],
      blockedParagraphIndices: [],
      separatorKind: 'leadingManualBreak',
      leadingManualBreakCount
    };
  }

  if (blankParagraphIndices.length === 0) {
    return {
      state: 'missing',
      lastAffiliationParagraphIndex: lastAffiliationParagraph.index,
      firstEmailParagraphIndex: firstEmailParagraph.index,
      blankParagraphIndices,
      removableBlankParagraphIndices: [],
      blockedParagraphIndices: [],
      separatorKind: 'none',
      leadingManualBreakCount
    };
  }

  const separatorKind = hasLeadingManualBreak ? 'mixed' : 'blankParagraph';
  const removableBlankParagraphIndices = hasLeadingManualBreak
    ? blankParagraphIndices
    : blankParagraphIndices.slice(1);
  return {
    state: blankParagraphIndices.length === 1 && !hasLeadingManualBreak ? 'passed' : 'extra',
    lastAffiliationParagraphIndex: lastAffiliationParagraph.index,
    firstEmailParagraphIndex: firstEmailParagraph.index,
    blankParagraphIndices,
    removableBlankParagraphIndices,
    blockedParagraphIndices: [],
    separatorKind,
    leadingManualBreakCount,
    reason: blankParagraphIndices.length === 1 && !hasLeadingManualBreak
      ? undefined
      : hasLeadingManualBreak
        ? 'A leading manual line break and one or more safe blank paragraphs create duplicate separation.'
        : `Expected one safe empty paragraph, found ${blankParagraphIndices.length}.`
  };
}

export interface FrontMatterTransition {
  fromIndex: number;
  toIndex: number;
  fromType: FrontMatterType;
  toType: FrontMatterType;
  blankIndices: number[];
  safeBlankIndices: number[];
  unsafeBlankIndices: number[];
  interveningNonBlankIndices: number[];
  expectedBlankCount: number;
  fromTextFingerprint: string;
  toTextFingerprint: string;
}

function frontMatterType(paragraph: ParagraphSnapshot): FrontMatterType | null {
  switch (paragraph.classification) {
    case 'title':
    case 'authorName':
    case 'authorAffiliation':
    case 'authorEmail':
    case 'abstract':
      return paragraph.classification;
    default:
      return null;
  }
}

/**
 * Analyze blank paragraphs between consecutive classified front-matter
 * components. Affiliation → email is intentionally returned with an
 * expected count of one for AFFILIATION_EMAIL_BLANK_001; the general extra
 * blank rule must skip that transition.
 */
export function analyzeFrontMatterTransitions(
  paragraphs: ParagraphSnapshot[]
): FrontMatterTransition[] {
  const ordered = [...paragraphs].sort((a, b) => a.index - b.index);
  const classified = ordered
    .map((paragraph, position) => ({ paragraph, position, type: frontMatterType(paragraph) }))
    .filter((entry): entry is { paragraph: ParagraphSnapshot; position: number; type: FrontMatterType } => entry.type !== null);
  const abstractEntry = classified.find(entry => entry.type === 'abstract');
  const frontMatter = abstractEntry
    ? classified.filter(entry => entry.position <= abstractEntry.position)
    : classified;
  const transitions: FrontMatterTransition[] = [];

  for (let index = 1; index < frontMatter.length; index += 1) {
    const previous = frontMatter[index - 1];
    const next = frontMatter[index];
    const between = ordered.slice(previous.position + 1, next.position);
    const blankParagraphs = between.filter(isBlankParagraph);
    const interveningNonBlankIndices = between
      .filter(paragraph => !isBlankParagraph(paragraph))
      .map(paragraph => paragraph.index);
    const blankIndices = blankParagraphs.map(paragraph => paragraph.index);
    const safeBlankIndices = blankParagraphs
      .filter(isStructurallySafeFrontMatterBlank)
      .map(paragraph => paragraph.index);
    const unsafeBlankIndices = blankParagraphs
      .filter(paragraph => !isStructurallySafeFrontMatterBlank(paragraph))
      .map(paragraph => paragraph.index);

    transitions.push({
      fromIndex: previous.paragraph.index,
      toIndex: next.paragraph.index,
      fromType: previous.type,
      toType: next.type,
      blankIndices,
      safeBlankIndices,
      unsafeBlankIndices,
      interveningNonBlankIndices,
      expectedBlankCount: previous.type === 'authorAffiliation' && next.type === 'authorEmail' ? 1 : 0,
      fromTextFingerprint: paragraphTextFingerprint(previous.paragraph.text),
      toTextFingerprint: paragraphTextFingerprint(next.paragraph.text)
    });
  }

  return transitions;
}

/**
 * Find figure captions whose next meaningful top-level body paragraph is not
 * separated by exactly one ordinary blank paragraph. Structural breaks and
 * non-body successors intentionally make the relationship inapplicable.
 */
export function analyzeFigureCaptionBodyGaps(paragraphs: ParagraphSnapshot[]): FigureCaptionBodyGapAnalysis {
  const issues: FigureCaptionBodyGapIssue[] = [];
  let applicableCaptionCount = 0;
  const ignoredCaptionIndices: number[] = [];
  const blockedCaptionIndices: number[] = [];
  paragraphs.forEach((caption, captionPosition) => {
    if (caption.classification !== 'figureCaption' || caption.tableNestingLevel !== 0) return;

    const blankParagraphIndices: number[] = [];
    let cursor = captionPosition + 1;
    while (cursor < paragraphs.length && isBlankParagraph(paragraphs[cursor])) {
      // A page/section/manual-break paragraph is not an ordinary separator;
      // do not insert or remove around it automatically.
      if (!isSafeTopLevelBlank(paragraphs[cursor])) {
        let afterStructural = cursor + 1;
        while (afterStructural < paragraphs.length && isBlankParagraph(paragraphs[afterStructural])) afterStructural += 1;
        const bodyAfterStructural = paragraphs[afterStructural];
        if (bodyAfterStructural?.classification === 'body'
          && bodyAfterStructural.tableNestingLevel === 0
          && !bodyAfterStructural.isListItem) {
          blockedCaptionIndices.push(caption.index);
          issues.push({
            captionParagraphIndex: caption.index,
            bodyParagraphIndex: bodyAfterStructural.index,
            blankParagraphIndices: [],
            blockedByStructuralParagraph: true
          });
        }
        return;
      }
      blankParagraphIndices.push(paragraphs[cursor].index);
      cursor += 1;
    }

    const next = paragraphs[cursor];
    if (!next
      || next.classification !== 'body'
      || next.tableNestingLevel !== 0
      || next.isListItem) {
      ignoredCaptionIndices.push(caption.index);
      return;
    }

    // A paragraph-level page break deliberately starts the body on a new
    // page. Do not insert an additional blank paragraph into that layout.
    if (next.hasPageBreakBefore === true) {
      blockedCaptionIndices.push(caption.index);
      issues.push({
        captionParagraphIndex: caption.index,
        bodyParagraphIndex: next.index,
        blankParagraphIndices: [],
        blockedByStructuralParagraph: true
      });
      return;
    }

    // A page or section break carried by the caption itself is structural.
    // Do not infer that an ordinary blank paragraph is needed after it.
    if (caption.blankParagraph?.hasExplicitPageBreak === true
      || caption.blankParagraph?.hasSectionBreak === true) {
      blockedCaptionIndices.push(caption.index);
      issues.push({
        captionParagraphIndex: caption.index,
        bodyParagraphIndex: next.index,
        blankParagraphIndices: [],
        blockedByStructuralParagraph: true
      });
      return;
    }

    applicableCaptionCount += 1;
    if (blankParagraphIndices.length !== 1) {
      issues.push({
        captionParagraphIndex: caption.index,
        bodyParagraphIndex: next.index,
        blankParagraphIndices
      });
    }
  });
  return {
    issues,
    applicableCaptionCount,
    ignoredCaptionIndices,
    blockedCaptionIndices
  };
}

export function findFigureCaptionBodyGapIssues(paragraphs: ParagraphSnapshot[]): FigureCaptionBodyGapIssue[] {
  return analyzeFigureCaptionBodyGaps(paragraphs).issues;
}

/**
 * Find removable blank paragraphs between a high-confidence external table
 * caption and its associated table block. Embedded table captions and any
 * intervening text/object are deliberately excluded from automatic cleanup.
 */
export function analyzeTableCaptionTableGaps(snapshot: DocumentSnapshot): TableCaptionTableGapAnalysis {
  const issues: TableCaptionTableGapIssue[] = [];
  const ignoredCaptionIndices: number[] = [];
  const blockedCaptionIndices: number[] = [];
  let applicableCaptionCount = 0;
  const paragraphsByIndex = new Map(snapshot.paragraphs.map(paragraph => [paragraph.index, paragraph]));
  const blocks = snapshot.layout.documentBlocks || [];

  for (const relation of snapshot.crossReferences?.tablePlacement || []) {
    const caption = paragraphsByIndex.get(relation.captionParagraphIndex);
    const table = relation.objectBlockIndex === null
      ? undefined
      : blocks.find(block => block.blockIndex === relation.objectBlockIndex);

    // Embedded captions have no external paragraph gap to normalize.
    if (!caption || caption.tableNestingLevel !== 0) {
      if (caption) ignoredCaptionIndices.push(caption.index);
      continue;
    }
    if (relation.confidence !== 'high' || relation.passed !== true || table?.kind !== 'table') {
      ignoredCaptionIndices.push(caption.index);
      continue;
    }

    const captionPosition = blocks.findIndex(block =>
      block.kind === 'paragraph' && block.paragraphIndex === caption.index
    );
    const tablePosition = blocks.findIndex(block => block.blockIndex === table.blockIndex);
    if (captionPosition < 0 || tablePosition <= captionPosition) {
      ignoredCaptionIndices.push(caption.index);
      continue;
    }

    const between = blocks.slice(captionPosition + 1, tablePosition);
    if (!between.length) {
      applicableCaptionCount += 1;
      continue;
    }

    // Any non-paragraph/nonblank content means this is not a simple gap. It
    // must not be deleted or silently treated as a caption-to-table spacing
    // issue.
    const betweenParagraphs = between.map(block =>
      block.kind === 'paragraph' && block.paragraphIndex !== undefined
        ? paragraphsByIndex.get(block.paragraphIndex)
        : undefined
    );
    if (betweenParagraphs.some(paragraph => !paragraph || !isBlankParagraph(paragraph))) {
      ignoredCaptionIndices.push(caption.index);
      continue;
    }

    const blankParagraphs = betweenParagraphs.filter(
      (paragraph): paragraph is ParagraphSnapshot => paragraph !== undefined
    );
    const blankParagraphIndices = blankParagraphs.map(paragraph => paragraph.index);
    if (blankParagraphs.every(isSafeTopLevelBlank)) {
      applicableCaptionCount += 1;
      issues.push({
        captionParagraphIndex: caption.index,
        tableBlockIndex: table.blockIndex,
        blankParagraphIndices
      });
      continue;
    }

    blockedCaptionIndices.push(caption.index);
    issues.push({
      captionParagraphIndex: caption.index,
      tableBlockIndex: table.blockIndex,
      blankParagraphIndices: [],
      blockedByStructuralParagraph: true
    });
  }

  return {
    issues,
    applicableCaptionCount,
    ignoredCaptionIndices,
    blockedCaptionIndices
  };
}

export function findTableCaptionTableGapIssues(snapshot: DocumentSnapshot): TableCaptionTableGapIssue[] {
  return analyzeTableCaptionTableGaps(snapshot).issues;
}

/**
 * Evaluate the single ordinary blank paragraph required immediately before a
 * confidently paired inline figure. Floating, table-contained, multi-drawing,
 * and otherwise ambiguous objects are deliberately not auto-evaluated.
 */
export function analyzeFigureSpaceAbove(snapshot: DocumentSnapshot): FigureSpaceAboveAnalysis {
  const issues: FigureSpaceAboveIssue[] = [];
  const ignoredCaptionIndices: number[] = [];
  const blockedCaptionIndices: number[] = [];
  let applicableFigureCount = 0;
  const paragraphsByIndex = new Map(snapshot.paragraphs.map(paragraph => [paragraph.index, paragraph]));
  const blocks = snapshot.layout.documentBlocks || [];

  for (const relation of snapshot.crossReferences?.figurePlacement || []) {
    const caption = paragraphsByIndex.get(relation.captionParagraphIndex);
    const object = relation.objectBlockIndex === null
      ? undefined
      : blocks.find(block => block.blockIndex === relation.objectBlockIndex);
    const figureParagraphIndex = object?.paragraphIndex;
    const inlineSingleFigure = relation.confidence === 'high'
      && object?.kind === 'paragraph'
      && object.hasInlineDrawing === true
      && object.hasAnchoredDrawing !== true
      && object.drawingCount === 1
      && figureParagraphIndex !== undefined;

    if (!caption || !inlineSingleFigure) {
      if (caption) blockedCaptionIndices.push(caption.index);
      continue;
    }

    const objectPosition = blocks.findIndex(block => block.blockIndex === object.blockIndex);
    if (objectPosition <= 0) {
      ignoredCaptionIndices.push(caption.index);
      continue;
    }

    const blankParagraphIndices: number[] = [];
    let cursor = objectPosition - 1;
    let structuralBlank = false;
    while (cursor >= 0) {
      const block = blocks[cursor];
      if (block.kind !== 'paragraph' || block.paragraphIndex === undefined) break;
      const paragraph = paragraphsByIndex.get(block.paragraphIndex);
      if (!paragraph || !isBlankParagraph(paragraph)) break;
      if (!isSafeTopLevelBlank(paragraph)) structuralBlank = true;
      else blankParagraphIndices.push(paragraph.index);
      cursor -= 1;
    }

    const previousBlock = blocks[cursor];
    const previous = previousBlock?.kind === 'paragraph' && previousBlock.paragraphIndex !== undefined
      ? paragraphsByIndex.get(previousBlock.paragraphIndex)
      : undefined;
    if (!previous || !previous.text.trim() || previous.tableNestingLevel !== 0 || previous.isListItem) {
      ignoredCaptionIndices.push(caption.index);
      continue;
    }
    if (structuralBlank) {
      blockedCaptionIndices.push(caption.index);
      issues.push({
        captionParagraphIndex: caption.index,
        figureParagraphIndex,
        previousParagraphIndex: previous.index,
        blankParagraphIndices: [],
        blockedByStructuralParagraph: true
      });
      continue;
    }

    applicableFigureCount += 1;
    if (blankParagraphIndices.length !== 1) {
      issues.push({
        captionParagraphIndex: caption.index,
        figureParagraphIndex,
        previousParagraphIndex: previous.index,
        blankParagraphIndices
      });
    }
  }

  return {
    issues,
    applicableFigureCount,
    ignoredCaptionIndices,
    blockedCaptionIndices
  };
}

export class ValidationEngine {
  static validate(strategy: StrategyDefinition, snapshot: DocumentSnapshot): ValidationSummary {
    const findings = strategy.rules.map(rule => this.evaluateRule(rule, snapshot));

    let complianceEarned = 0;
    let compliancePossible = 0;
    let overallEarned = 0;
    let overallPossible = 0;
    for (const finding of findings) {
      if (finding.status !== 'passed' && finding.status !== 'failed') continue;
      const w = severityWeight[finding.severity];
      overallPossible += w;
      if (finding.status === 'passed') overallEarned += w;
      if (finding.assessmentRole === 'consistency') continue;
      compliancePossible += w;
      if (finding.status === 'passed') complianceEarned += w;
    }

    const roundScore = (earned: number, possible: number): number =>
      possible > 0 ? Math.round((earned / possible) * 1000) / 10 : 0;
    const complianceScore = roundScore(complianceEarned, compliancePossible);
    const overallScore = roundScore(overallEarned, overallPossible);
    const compliancePassed = findings.filter(f => f.assessmentRole !== 'consistency' && f.status === 'passed').length;
    const complianceFailed = findings.filter(f => f.assessmentRole !== 'consistency' && f.status === 'failed').length;

    return {
      findings,
      score: complianceScore,
      complianceScore,
      overallScore,
      passed: compliancePassed,
      failed: complianceFailed,
      compliancePassed,
      complianceFailed,
      consistencyPassed: findings.filter(f => f.assessmentRole === 'consistency' && f.status === 'passed').length,
      consistencyFailed: findings.filter(f => f.assessmentRole === 'consistency' && f.status === 'failed').length,
      unresolvedFailures: findings.filter(f => f.status === 'failed').length,
      fixableFailures: findings.filter(f => f.status === 'failed' && f.autoFixable === true).length,
      manualReview: findings.filter(f => f.status === 'manual_review').length,
      checklist: findings.filter(f => f.status === 'checklist').length,
      notApplicable: findings.filter(f => f.status === 'not_applicable').length,
      notEvaluated: findings.filter(f => f.status === 'not_evaluated').length,
      scoredRules: findings.filter(f => f.assessmentRole !== 'consistency' && (f.status === 'passed' || f.status === 'failed')).length
    };
  }

  private static evaluateRule(rule: RuleDefinition, snapshot: DocumentSnapshot): Finding {
    if (rule.conflict?.length) {
      const alternatives = rule.conflict.map(c => `${c.source}: ${displayValue(c.value, rule.unit)}`).join(' · ');
      return this.finding(
        rule,
        'manual_review',
        [],
        this.expected(rule),
        alternatives,
        'The source materials conflict, so FormaScholar does not silently choose one requirement.',
        undefined,
        'source_conflict'
      );
    }

    const requiresParagraphMapping = rule.target === 'body'
      || rule.id === 'LINE_SPACING_001'
      || rule.id === 'TEXT_MULTIPLE_SPACES_001'
      || rule.id === 'CONCLUSION_BODY_FONT_001';
    if (requiresParagraphMapping && snapshot.layout.paragraphMappingAligned === false) {
      return this.finding(
        rule,
        'not_evaluated',
        [],
        this.expected(rule),
        'Paragraph/layout mapping unavailable',
        'The live Word paragraph sequence and OOXML paragraph sequence do not align, so this rule was not scored safely.',
        undefined,
        'technical'
      );
    }

    // Determine paragraph targets before deciding that an unimplemented rule is
    // a manual review item. If the manuscript has no such object, the correct
    // status is not applicable rather than a noisy not-evaluated card.
    const rawTargets = rule.target === 'document'
      ? []
      : snapshot.paragraphs.filter(p => p.classification === rule.target);
    // A bullet/list paragraph's negative first-line indent is part of its
    // hanging-list layout, not a violation of the manuscript body indent.
    // Keep list typography and other body checks active; exclude only this
    // normal-body indentation rule.
    let targets = rule.target === 'body'
      ? rawTargets.filter(p => isReliableLayoutParagraph(p) && (rule.id !== 'BODY_FIRST_LINE_INDENT_001' || !p.isListItem))
      : rule.id === 'BODY_FIRST_LINE_INDENT_001'
        ? rawTargets.filter(p => !p.isListItem)
        : rawTargets;
    if (rule.id === 'HEADING1_SPACE_BEFORE_001') {
      const firstHeading = findFirstHeading1AfterAbstract(snapshot.paragraphs);
      if (firstHeading) targets = targets.filter(paragraph => paragraph.index !== firstHeading.index);
    }

    if (rule.reviewMode || rule.appliesWhen) {
      const applicability = ApplicabilityAnalyzer.check(rule, snapshot);
      if (!applicability.applies) {
        return this.finding(rule, 'not_applicable', [], this.expected(rule), 'Not applicable', applicability.reason);
      }
    }

    const unreliableTargets = rawTargets.filter(p => p.ooxmlAligned === false);
    if (requiresParagraphMapping && unreliableTargets.length) {
      return this.finding(
        rule,
        'not_evaluated',
        unreliableTargets.map(p => p.index),
        this.expected(rule),
        'Paragraph mapping unavailable',
        'Matching content exists, but its live Word/OOXML paragraph mapping could not be verified safely.',
        undefined,
        'technical'
      );
    }

    if (rule.target !== 'document' && rule.property !== 'presence' && targets.length === 0) {
      return this.finding(rule, 'not_applicable', [], this.expected(rule), 'No matching content',
        `No ${rule.target} content was detected, so this rule does not apply to this manuscript.`);
    }

    if (!rule.implemented) {
      return this.evaluateUnimplementedRule(rule, targets);
    }

    if (rule.target === 'document') return this.evaluateDocumentRule(rule, snapshot);

    if (rule.property === 'presence') {
      const exists = targets.length > 0;
      return this.finding(
        rule,
        exists ? 'passed' : 'failed',
        targets.map(p => p.index),
        'Present',
        exists ? 'Present' : 'Missing',
        exists ? `${rule.name} detected.` : `${rule.name} was not detected.`
      );
    }

    if (rule.property.startsWith('crossref.')
      || rule.property === 'structure.captionBelowFigure'
      || rule.property === 'structure.captionAboveTable'
      || rule.property === 'structure.figurePriorCitation') {
      return this.evaluateCrossReferenceRule(rule, snapshot);
    }

    if (rule.property === 'structure.figureSpaceAbove') {
      return this.evaluateFigureSpaceAbove(rule, snapshot);
    }

    if (rule.property === 'structure.tableCaptionTableGap') {
      return this.evaluateTableCaptionTableGap(rule, snapshot);
    }

    if (rule.property === 'paragraph.frontMatterSpacing') {
      const finding = this.evaluateFrontMatterSpacing(rule, targets);
      if (finding.status === 'failed') {
        finding.fixParagraphTexts = Object.fromEntries(
          targets
            .filter(paragraph => finding.paragraphIndices.includes(paragraph.index))
            .map(paragraph => [paragraph.index, paragraph.text])
        );
      }
      return finding;
    }

    if (rule.property.startsWith('text.')
      || rule.property.startsWith('structure.')
      || rule.property.startsWith('run.')
      || rule.property.startsWith('reference.')) {
      return this.evaluateTextOrStructure(rule, targets, snapshot);
    }

    const failures: ParagraphSnapshot[] = [];
    const unavailable: ParagraphSnapshot[] = [];
    const detectedValues: string[] = [];
    for (const p of targets) {
      const requiresRunLevelItalic = rule.id === 'AUTHOR_EMAIL_ITALIC_001'
        || rule.id === 'AUTHOR_AFFILIATION_ITALIC_001';

      if (requiresRunLevelItalic
        && p.authorAssociationBodyItalic !== null
        && p.authorAssociationBodyItalic !== undefined) {
        const detected = p.authorAssociationBodyItalic;
        detectedValues.push(displayValue(detected, rule.unit));
        if (!this.matches(detected, rule.expected, rule.property)) failures.push(p);
        continue;
      }

      // For ordinary typography, a non-null live Word paragraph value is the
      // freshest representation after an Office.js correction. OOXML runs are
      // the fallback for mixed/unavailable paragraph properties. Association
      // italics remain run-level because their leading markers intentionally
      // have different formatting from the text that follows them.
      const liveDetected = requiresRunLevelItalic
        ? null
        : this.paragraphProperty(p, rule.property);
      if (liveDetected !== null && liveDetected !== undefined) {
        detectedValues.push(displayValue(liveDetected, rule.unit));
        if (!this.matches(liveDetected, rule.expected, rule.property)) failures.push(p);
        continue;
      }

      const runValues = this.runPropertyValues(p, rule.property);
      if (runValues) {
        let hasUnavailableRun = false;
        let hasFailure = false;
        for (const value of runValues) {
          if (value === null || value === undefined) {
            hasUnavailableRun = true;
            continue;
          }
          detectedValues.push(displayValue(value, rule.unit));
          if (!this.matches(value, rule.expected, rule.property)) hasFailure = true;
        }
        if (hasFailure) failures.push(p);
        if (hasUnavailableRun) unavailable.push(p);
        continue;
      }

      const detected = this.paragraphProperty(p, rule.property);
      if (detected === null || detected === undefined) {
        unavailable.push(p);
        continue;
      }
      if (!this.matches(detected, rule.expected, rule.property)) failures.push(p);
      detectedValues.push(displayValue(detected, rule.unit));
    }

    const unique = [...new Set(detectedValues)];
    const detected = unique.length <= 3 ? unique.join(', ') : `${unique.slice(0, 3).join(', ')} + ${unique.length - 3} more`;

    if (failures.length) {
      const associationItalicFailure = rule.id === 'AUTHOR_EMAIL_ITALIC_001'
        || rule.id === 'AUTHOR_AFFILIATION_ITALIC_001';
      const associationLabel = rule.id === 'AUTHOR_EMAIL_ITALIC_001'
        ? 'Email'
        : 'Affiliation';
      const expectsItalic = rule.expected === true;
      const oppositeItalicDescription = expectsItalic ? 'non-italic' : 'italic';
      const associationItalicDetected =
        `${associationLabel} body contains ${oppositeItalicDescription} text; association markers are excluded`;
      const associationItalicUnavailable =
        `${failures.length} ${associationLabel.toLowerCase()} paragraph(s) contain ${oppositeItalicDescription} ${associationLabel.toLowerCase()} text; association markers such as a), b), and c) are excluded. ${unavailable.length} paragraph(s) could not be evaluated.`;
      const associationItalicMessage =
        `${failures.length} ${associationLabel.toLowerCase()} paragraph(s) contain ${oppositeItalicDescription} ${associationLabel.toLowerCase()} text. Leading association markers are excluded from this rule.`;
      return this.finding(
        rule,
        'failed',
        failures.map(p => p.index),
        this.expected(rule),
        associationItalicFailure
          ? associationItalicDetected
          : unavailable.length ? `${failures.length} failed; ${unavailable.length} unavailable` : detected,
        unavailable.length
          ? associationItalicFailure
            ? associationItalicUnavailable
            : `${failures.length} paragraph(s) definitely fail; ${unavailable.length} could not be evaluated.`
          : associationItalicFailure
            ? associationItalicMessage
            : `${failures.length} of ${targets.length} matching paragraph${targets.length === 1 ? '' : 's'} do not comply.`
      );
    }

    if (unavailable.length) {
      return this.finding(
        rule,
        'not_evaluated',
        unavailable.map(p => p.index),
        this.expected(rule),
        `${unavailable.length} paragraph(s) unavailable${detected ? `; detected ${detected}` : ''}`,
        `${rule.name} could not be evaluated for ${unavailable.length} matching paragraph(s) because effective formatting was unavailable.`,
        undefined,
        'technical'
      );
    }

    const status = failures.length === 0 ? 'passed' : 'failed';
    const message = failures.length === 0
      ? `${targets.length} matching paragraph${targets.length === 1 ? '' : 's'} comply.`
      : `${failures.length} of ${targets.length} matching paragraph${targets.length === 1 ? '' : 's'} do not comply.`;

    return this.finding(rule, status, failures.map(p => p.index), this.expected(rule), detected, message);
  }

  private static evaluateUnimplementedRule(rule: RuleDefinition, targets: ParagraphSnapshot[]): Finding {
    const mode = rule.reviewMode
      || (rule.category === 'Submission' ? 'submission' : undefined)
      || (rule.property.startsWith('manual.') ? 'manual' : 'technical');

    if (mode === 'submission') {
      return this.finding(
        rule,
        'checklist',
        targets.map(p => p.index),
        this.expected(rule),
        'User confirmation required',
        rule.note || 'This requirement concerns submission preparation or evidence outside the open DOCX and is presented as a checklist item.',
        undefined,
        'submission'
      );
    }

    if (mode === 'manual') {
      return this.finding(
        rule,
        'manual_review',
        targets.map(p => p.index),
        this.expected(rule),
        'Manual review required',
        rule.note || 'This requirement needs human judgement and is not included in the automated compliance score.',
        undefined,
        'manual'
      );
    }

    return this.finding(
      rule,
      'not_evaluated',
      targets.map(p => p.index),
      this.expected(rule),
      'Not evaluated',
      rule.note || 'This deterministic check is not yet available in the current Word-web implementation.',
      undefined,
      'technical'
    );
  }

  private static evaluateDocumentRule(rule: RuleDefinition, snapshot: DocumentSnapshot): Finding {
    if (rule.property.startsWith('layout.')) {
      if (!snapshot.layout.ooxmlAvailable) {
        return this.finding(rule, 'not_evaluated', [], this.expected(rule), 'OOXML unavailable',
          snapshot.layout.ooxmlError || 'The compressed Word document could not be inspected.', undefined, 'technical');
      }
      const key = rule.property.substring('layout.'.length) as keyof typeof snapshot.layout;
      const sectionKey = rule.property.substring('layout.'.length) as keyof NonNullable<typeof snapshot.layout.sections>[number];
      const sectionProperties = new Set([
        'pageWidthInches', 'pageHeightInches', 'marginTopInches', 'marginRightInches',
        'marginBottomInches', 'marginLeftInches', 'columns'
      ]);
      if (sectionProperties.has(String(sectionKey)) && snapshot.layout.sections?.length) {
        const sections = snapshot.layout.sections;
        const invalid = sections.filter(section => {
          const value = section[sectionKey];
          if (value === undefined || value === null) return true;
          if (sectionKey === 'pageWidthInches' || sectionKey === 'pageHeightInches') {
            const landscape = section.orientation === 'landscape'
              || (section.pageWidthInches !== undefined && section.pageHeightInches !== undefined
                && section.pageWidthInches > section.pageHeightInches);
            const expected = sectionKey === 'pageWidthInches'
              ? landscape && Number(rule.expected) === 8.5 ? 11 : Number(rule.expected)
              : landscape && Number(rule.expected) === 11 ? 8.5 : Number(rule.expected);
            return !this.matches(value, expected, rule.property);
          }
          return !this.matches(value, rule.expected, rule.property);
        });
        const formatSection = (section: typeof sections[number]) => {
          const value = section[sectionKey];
          return `Section ${section.index + 1} = ${displayValue(value, rule.unit)}`;
        };
        if (invalid.some(section => section[sectionKey] === undefined || section[sectionKey] === null)) {
          return this.finding(
            rule,
            'not_evaluated',
            [],
            this.expected(rule),
            invalid.map(formatSection).join('; '),
            `${rule.name} could not be evaluated for every document section.`,
            undefined,
            'technical'
          );
        }
        const pass = invalid.length === 0;
        return this.finding(
          rule,
          pass ? 'passed' : 'failed',
          [],
          this.expected(rule),
          pass
            ? sections.length === 1
              ? displayValue(sections[0][sectionKey], rule.unit)
              : `All ${sections.length} sections = ${displayValue(sections[0][sectionKey], rule.unit)}`
            : invalid.map(formatSection).join('; '),
          pass ? `${rule.name} complies in every document section.` : `${rule.name} does not match in every document section.`
        );
      }
      const detected = snapshot.layout[key];
      if (detected === undefined || detected === null) {
        return this.finding(rule, 'not_evaluated', [], this.expected(rule), 'OOXML value unavailable',
          `${rule.name} could not be read from the document section properties, so it was not scored.`, undefined, 'technical');
      }
      const pass = this.matches(detected, rule.expected, rule.property);
      return this.finding(rule, pass ? 'passed' : 'failed', [], this.expected(rule), displayValue(detected, rule.unit),
        pass ? `${rule.name} complies.` : `${rule.name} does not match the configured AIP strategy.`);
    }

    const requiredSectionGroup = REQUIRED_SECTION_GROUPS[rule.property];
    if (requiredSectionGroup) {
      return this.evaluateRequiredSectionGroup(rule, snapshot, requiredSectionGroup);
    }

    switch (rule.property) {
      case 'document.abstractPresence': {
        const abstract = snapshot.paragraphs.find(p => p.classification === 'abstract');
        return this.finding(
          rule,
          abstract ? 'passed' : 'failed',
          abstract ? [abstract.index] : [],
          this.expected(rule),
          abstract ? 'Abstract present' : 'Abstract missing',
          abstract
            ? 'The manuscript contains an Abstract section.'
            : 'The AIP template states that each paper must include an abstract.'
        );
      }
      case 'structure.firstHeading1SpaceBefore': {
        const firstHeading = findFirstHeading1AfterAbstract(snapshot.paragraphs);
        if (!firstHeading) {
          return this.finding(
            rule,
            'not_applicable',
            [],
            this.expected(rule),
            'No Heading 1 follows the Abstract',
            'The first-post-Abstract Heading 1 exception does not apply because an Abstract or following Heading 1 was not detected.'
          );
        }
        if (firstHeading.tableNestingLevel !== 0
          || firstHeading.layoutSensitive === true
          || firstHeading.ooxmlAligned === false
          || snapshot.layout.ooxmlAvailable === false
          || snapshot.layout.paragraphMappingAligned === false) {
          return this.finding(
            rule,
            'not_evaluated',
            [firstHeading.index],
            this.expected(rule),
            'Heading/layout mapping unavailable',
            'The first Heading 1 after the Abstract is inside a table, positioned layout, or unreliable paragraph mapping, so its spacing was not scored safely.',
            undefined,
            'technical'
          );
        }
        if (firstHeading.spaceBefore === null || firstHeading.spaceBefore === undefined) {
          return this.finding(
            rule,
            'not_evaluated',
            [firstHeading.index],
            this.expected(rule),
            'Space-before unavailable',
            'The first Heading 1 after the Abstract does not expose a reliable paragraph spacing value.',
            undefined,
            'technical'
          );
        }
        const passed = approxEqual(firstHeading.spaceBefore, 0, 0.75);
        const finding = this.finding(
          rule,
          passed ? 'passed' : 'failed',
          [firstHeading.index],
          this.expected(rule),
          `${firstHeading.spaceBefore} pt`,
          passed
            ? 'The first Heading 1 after the Abstract uses 0 pt space before.'
            : 'The first Heading 1 after the Abstract should use 0 pt space before.'
        );
        if (!passed) {
          finding.fixParagraphTexts = { [firstHeading.index]: firstHeading.text };
        }
        return finding;
      }
      case 'structure.affiliationEmailBlank': {
        const analysis = analyzeAffiliationEmailBlankLine(snapshot.paragraphs);
        const contextIndices = [
          analysis.lastAffiliationParagraphIndex,
          analysis.firstEmailParagraphIndex,
          ...analysis.blankParagraphIndices
        ].filter((index): index is number => index !== undefined);
        if (analysis.state === 'not_applicable') {
          return this.finding(
            rule,
            'not_applicable',
            contextIndices,
            this.expected(rule),
            'Affiliation/email separator not applicable',
            analysis.reason || 'A last affiliation followed by an author email was not detected.'
          );
        }
        if (analysis.state === 'blocked') {
          return this.finding(
            rule,
            'not_evaluated',
            [...contextIndices, ...analysis.blockedParagraphIndices],
            this.expected(rule),
            'Separator contains protected structural content',
            analysis.reason || 'The separator could not be evaluated safely.',
            undefined,
            'technical'
          );
        }
        if (analysis.state === 'passed') {
          const leadingBreakOnly = analysis.separatorKind === 'leadingManualBreak';
          return this.finding(
            rule,
            'passed',
            contextIndices,
            this.expected(rule),
            leadingBreakOnly
              ? '1 leading manual line break — compliant'
              : '1 safe blank paragraph — compliant',
            leadingBreakOnly
              ? 'One leading manual line break separates the last affiliation from the first author email.'
              : 'Exactly one safe blank paragraph separates the last affiliation from the first author email.'
          );
        }
        const count = analysis.blankParagraphIndices.length;
        if (analysis.state === 'missing') {
          return this.finding(
            rule,
            'failed',
            contextIndices,
            this.expected(rule),
            'No affiliation-email separator detected',
            'No safe blank paragraph or leading manual line break separates the last affiliation from the first author email; one safe blank paragraph will be inserted.'
          );
        }
        return this.finding(
          rule,
          'failed',
          contextIndices,
          this.expected(rule),
          analysis.separatorKind === 'mixed'
            ? 'Duplicate separator: blank paragraph and leading manual break'
            : `${count} safe empty paragraphs`,
          analysis.separatorKind === 'mixed'
            ? 'A leading manual line break and a safe blank paragraph create duplicate separation; only the redundant blank paragraph will be removed.'
            : `${count} safe blank paragraphs separate the last affiliation from the first author email; the extras can be removed safely.`
        );
      }
      case 'structure.frontMatterBlankStructure': {
        const transitions = analyzeFrontMatterTransitions(snapshot.paragraphs);
        const applicable = transitions.filter(transition =>
          transition.expectedBlankCount === 0
          && transition.interveningNonBlankIndices.length === 0
          && !(transition.fromType === 'authorAffiliation' && transition.toType === 'authorEmail')
        );
        if (!applicable.length) {
          return this.finding(
            rule,
            'not_applicable',
            [],
            this.expected(rule),
            'No applicable front-matter zero-blank transitions',
            'No consecutive front-matter components requiring zero blank paragraphs were detected.'
          );
        }

        const blocked = applicable.filter(transition => transition.unsafeBlankIndices.length > 0);
        const offenders = applicable.filter(transition =>
          transition.blankIndices.length > 0
          && transition.unsafeBlankIndices.length === 0
        );
        const safeBlankIndices = [...new Set(
          offenders.flatMap(transition => transition.safeBlankIndices)
        )];
        const affectedIndices = [...new Set([
          ...offenders.flatMap(transition => [transition.fromIndex, transition.toIndex, ...transition.blankIndices]),
          ...blocked.flatMap(transition => [transition.fromIndex, transition.toIndex, ...transition.unsafeBlankIndices])
        ])];

        if (offenders.length) {
          const details = offenders
            .map(transition => `${transition.fromType} → ${transition.toType}: ${transition.blankIndices.length}`)
            .join('; ');
          const finding = this.finding(
            rule,
            'failed',
            affectedIndices,
            this.expected(rule),
            `${offenders.length} transition${offenders.length === 1 ? '' : 's'} contain extra blank paragraphs (${details})`,
            blocked.length
              ? `${safeBlankIndices.length} safe front-matter blank paragraph(s) can be removed; ${blocked.length} transition${blocked.length === 1 ? '' : 's'} require manual review because protected blank content was detected.`
              : `${offenders.length} front-matter transition${offenders.length === 1 ? '' : 's'} contain unnecessary blank paragraphs.`,
            safeBlankIndices,
            undefined,
            undefined,
            safeBlankIndices.length > 0
          );
          finding.frontMatterFixTargets = offenders.map<FrontMatterTransitionFixTarget>(transition => ({
            fromType: transition.fromType,
            toType: transition.toType,
            fromIndex: transition.fromIndex,
            toIndex: transition.toIndex,
            fromTextFingerprint: transition.fromTextFingerprint,
            toTextFingerprint: transition.toTextFingerprint,
            blankIndices: [...transition.blankIndices],
            expectedBlankCount: transition.expectedBlankCount
          }));
          return finding;
        }

        if (blocked.length) {
          return this.finding(
            rule,
            'not_evaluated',
            blocked.flatMap(transition => [transition.fromIndex, transition.toIndex, ...transition.unsafeBlankIndices]),
            this.expected(rule),
            'Protected blank paragraph in front matter',
            'One or more front-matter transitions contain a structural or positioned blank paragraph that cannot be removed automatically.',
            undefined,
            'technical'
          );
        }

        return this.finding(
          rule,
          'passed',
          [],
          this.expected(rule),
          'No extra front-matter blank paragraphs',
          'All applicable zero-blank front-matter transitions comply.'
        );
      }
      case 'structure.heading1Style': {
        const candidates = snapshot.paragraphs.filter(isSafeHeading1StyleCandidate);
        if (!candidates.length) {
          return this.finding(
            rule,
            'not_applicable',
            [],
            this.expected(rule),
            'No confidently identified first-level heading candidates',
            'No safe first-level section-heading candidates were detected for style correction.'
          );
        }
        const failures = candidates.filter(paragraph => !hasHeading1ParagraphStyle(paragraph));
        const finding = this.finding(
          rule,
          failures.length ? 'failed' : 'passed',
          failures.map(paragraph => paragraph.index),
          this.expected(rule),
          failures.length
            ? failures.map(paragraph => `${paragraph.text.trim()}: ${paragraph.style || 'Unknown'}`).join(' · ')
            : 'Heading 1 style',
          failures.length
            ? `${failures.length} confidently identified first-level heading${failures.length === 1 ? '' : 's'} use a different Word paragraph style.`
            : `${candidates.length} first-level heading${candidates.length === 1 ? '' : 's'} use the Heading 1 paragraph style.`
        );
        if (failures.length) {
          finding.fixParagraphTexts = Object.fromEntries(
            failures.map(paragraph => [paragraph.index, paragraph.text])
          );
        }
        return finding;
      }
      case 'structure.referenceSectionConsistency': {
        const referenceCitations = snapshot.crossReferences?.referenceCitations || [];
        // invalidReferenceCitations is a filtered subset of referenceCitations,
        // not a second citation population. Count each in-text citation once.
        const citationCount = referenceCitations.length;
        const referenceHeading = snapshot.paragraphs.find(p =>
          p.classification === 'referencesHeading'
          || p.sectionFamilies?.includes('references')
        );
        if (citationCount === 0) {
          return this.finding(
            rule,
            'not_applicable',
            [],
            this.expected(rule),
            'No in-text reference citations detected',
            'A References/Bibliography section is checked when in-text citations are present; no such citations were detected.'
          );
        }
        const passed = referenceHeading !== undefined;
        return this.finding(
          rule,
          passed ? 'passed' : 'failed',
          passed ? [referenceHeading.index] : [],
          this.expected(rule),
          passed ? 'References/Bibliography present' : 'References/Bibliography missing',
          passed
            ? `${citationCount} in-text citation${citationCount === 1 ? '' : 's'} have a detected References/Bibliography section.`
            : `${citationCount} in-text citation${citationCount === 1 ? '' : 's'} were detected, but no References/Bibliography section could be identified.`
        );
      }
      case 'structure.sectionStructure': {
        const sectionParagraphs = snapshot.paragraphs.filter(isRecognizedFirstLevelSection);
        if (!sectionParagraphs.length) {
          return this.finding(
            rule,
            'not_applicable',
            [],
            this.expected(rule),
            'No confidently recognized first-level sections',
            'No confidently recognizable section titles were found for the advisory structure review.'
          );
        }
        const titles = [...new Set(sectionParagraphs.map(p => p.sectionNormalizedTitle || p.text.trim()))];
        const families = [...new Set(sectionParagraphs.flatMap(p => p.sectionFamilies || []))];
        const familyText = families.length
          ? families.map(sectionFamilyLabel).join(', ')
          : 'custom first-level headings';
        return this.finding(
          rule,
          'manual_review',
          sectionParagraphs.map(p => p.index),
          'Advisory section-family recognition; no fixed IMRaD sequence required',
          titles.join(' · '),
          `Recognized ${sectionParagraphs.length} first-level section heading${sectionParagraphs.length === 1 ? '' : 's'} (${familyText}). AIP section names are flexible; review the manuscript structure manually rather than treating absent conventional sections as compliance failures.`,
          undefined,
          'manual'
        );
      }
      case 'document.singleLineSpacing':
        return this.evaluateSingleLineSpacing(rule, snapshot);
      case 'document.multipleSpaces': {
        const eligible = snapshot.paragraphs.filter(p =>
          p.tableNestingLevel === 0
          && p.layoutSensitive !== true
          && p.classification !== 'figureCaption'
          && p.classification !== 'tableCaption'
          && p.classification !== 'equation'
          && p.text.trim().length > 0
        );
        const unreliable = eligible.filter(p => p.ooxmlAligned === false);
        if (unreliable.length) {
          return this.finding(
            rule,
            'not_evaluated',
            unreliable.map(p => p.index),
            this.expected(rule),
            'Paragraph mapping unavailable',
            'Repeated-space analysis was not scored because one or more eligible paragraphs could not be mapped safely to OOXML.',
            undefined,
            'technical'
          );
        }
        const offenders = eligible.filter(p =>
          p.ooxmlAligned !== false
          && /\S {2,}\S/.test(p.text)
        );
        return this.booleanDocumentFinding(
          rule,
          offenders.length === 0,
          offenders.map(p => p.index),
          offenders.length ? `${offenders.length} paragraph(s) contain repeated internal spaces` : 'No repeated internal spaces detected',
          'No repeated internal spaces were detected in eligible manuscript prose.',
          'Repeated internal spaces were detected in eligible manuscript prose.'
        );
      }
      case 'structure.extraBlankParagraphs': {
        if (!snapshot.layout.ooxmlAvailable) {
          return this.finding(
            rule,
            'not_evaluated',
            [],
            this.expected(rule),
            'OOXML unavailable',
            snapshot.layout.ooxmlError || 'Blank paragraph structure could not be inspected.',
            undefined,
            'technical'
          );
        }
        const offenders = findExtraBlankBodyParagraphIndices(snapshot.paragraphs);
        return this.booleanDocumentFinding(
          rule,
          offenders.length === 0,
          offenders,
          offenders.length ? `${offenders.length} unnecessary blank paragraph(s) detected` : 'No unnecessary blank paragraphs detected',
          'No safely removable blank paragraphs were detected between body paragraphs.',
          'Safely removable blank paragraphs were detected between body paragraphs.'
        );
      }
      case 'structure.figureCaptionPresence': {
        const features = snapshot.features || ApplicabilityAnalyzer.analyze(snapshot);
        const figureObjects = (snapshot.layout.documentBlocks || [])
          .filter(block => (block.drawingCount || 0) > 0);
        const captionedObjects = new Set(
          (snapshot.crossReferences?.figurePlacement || [])
            .filter(relation => relation.objectBlockIndex !== null)
            .map(relation => relation.objectBlockIndex as number)
        );
        const missingObjects = figureObjects.filter(block => !captionedObjects.has(block.blockIndex));
        const drawingParagraphs = (snapshot.layout.documentBlocks || [])
          .filter(block => missingObjects.some(missing => missing.blockIndex === block.blockIndex) && block.paragraphIndex !== undefined)
          .map(block => block.paragraphIndex as number);
        return this.finding(
          rule,
          missingObjects.length === 0 ? 'passed' : 'failed',
          missingObjects.length === 0 ? [] : drawingParagraphs,
          this.expected(rule),
          missingObjects.length === 0 ? 'Every detected figure candidate has a caption' : `${missingObjects.length} figure/drawing candidate(s) have no detected caption`,
          missingObjects.length === 0
            ? 'Figure caption coverage is complete.'
            : 'A drawing candidate was detected without a corresponding FIGURE caption. Review whether it is a manuscript figure and add a numbered caption if required.'
        );
      }
      case 'structure.tableCaptionPresence': {
        const features = snapshot.features || ApplicabilityAnalyzer.analyze(snapshot);
        const tableObjects = (snapshot.layout.documentBlocks || [])
          .filter(block => block.kind === 'table'
            && block.isEquationLayout !== true
            && (block.drawingCount || 0) === 0);
        const captionedObjects = new Set(
          (snapshot.crossReferences?.tablePlacement || [])
            .filter(relation => relation.objectBlockIndex !== null)
            .map(relation => relation.objectBlockIndex as number)
        );
        const missingObjects = tableObjects.filter(block => !captionedObjects.has(block.blockIndex));
        return this.finding(
          rule,
          missingObjects.length === 0 ? 'passed' : 'failed',
          [],
          this.expected(rule),
          missingObjects.length === 0 ? 'Every detected table candidate has a caption' : `${missingObjects.length} table candidate(s) have no detected caption`,
          missingObjects.length === 0
            ? 'Table caption coverage is complete.'
            : 'A table candidate was detected without a corresponding TABLE caption. Review whether it is a manuscript table and add a numbered caption if required.'
        );
      }
      case 'document.guidanceNotationRemoved': {
        const offenders = snapshot.paragraphs.filter(p => /\((?:First|Second|Third)\s+Level\s+Heading\)/i.test(p.text));
        return this.booleanDocumentFinding(
          rule,
          offenders.length === 0,
          offenders.map(p => p.index),
          offenders.length ? `${offenders.length} template guidance notation paragraph(s) detected` : 'No level-heading guidance notation detected',
          'Template level-heading guidance notation has been removed.',
          'Template guidance notation remains in the manuscript.'
        );
      }
      case 'document.styleInstructionRemoved': {
        const offenders = snapshot.paragraphs.filter(p =>
          /Use\s+the\s+Microsoft\s+Word\s+template\s+style\s*:/i.test(p.text)
          || /Use\s+Times\s+New\s+Roman\s+Font\s*:/i.test(p.text)
        );
        return this.booleanDocumentFinding(
          rule,
          offenders.length === 0,
          offenders.map(p => p.index),
          offenders.length ? `${offenders.length} template instruction paragraph(s) detected` : 'No template style-instruction text detected',
          'Template style-instruction text has been removed.',
          'Template style-instruction text remains in the manuscript.'
        );
      }
      case 'document.templateBoilerplateRemoved': {
        const boilerplate = [
          'This document was prepared using the AIP Conference Proceedings template',
          'How to Apply the AIP Conference Proceedings Template Paragraph Styles',
          'How to Use this Template',
          'Click here to display the Styles window',
          'To apply a style to a paragraph',
          'Where to Find Further Information',
          'Summary: Points to Consider when Preparing Your Paper'
        ].map(value => this.normalizeTemplateText(value));
        const offenders = snapshot.paragraphs.filter(p => {
          const text = this.normalizeTemplateText(p.text);
          return boilerplate.some(phrase => text.includes(phrase));
        });
        return this.booleanDocumentFinding(
          rule,
          offenders.length === 0,
          offenders.map(p => p.index),
          offenders.length ? `${offenders.length} template instruction paragraph(s) detected` : 'No high-confidence AIP template boilerplate detected',
          'High-confidence AIP template boilerplate has been removed.',
          'High-confidence AIP template boilerplate remains in the manuscript.'
        );
      }
      case 'document.conclusionTypography':
        return this.evaluateSectionBodyTypography(rule, snapshot, 'CONCLUSION');
      case 'document.acknowledgmentsTypography':
        return this.evaluateSectionBodyTypography(rule, snapshot, 'ACKNOWLEDGMENTS');
      default:
        return this.finding(rule, 'not_evaluated', [], this.expected(rule), 'Not evaluated',
          rule.note || 'This document-level rule is not yet implemented.', undefined, 'technical');
    }
  }

  private static evaluateFrontMatterSpacing(
    rule: RuleDefinition,
    targets: ParagraphSnapshot[]
  ): Finding {
    const expected = Number(rule.expected);
    const failures: ParagraphSnapshot[] = [];
    const unavailable: ParagraphSnapshot[] = [];
    const detected: string[] = [];

    for (const paragraph of targets) {
      if (paragraph.spaceBefore === null
        || paragraph.spaceBefore === undefined
        || paragraph.spaceAfter === null
        || paragraph.spaceAfter === undefined) {
        unavailable.push(paragraph);
        continue;
      }
      detected.push(`before ${paragraph.spaceBefore} pt / after ${paragraph.spaceAfter} pt`);
      if (!approxEqual(paragraph.spaceBefore, expected, 0.75)
        || !approxEqual(paragraph.spaceAfter, expected, 0.75)) {
        failures.push(paragraph);
      }
    }

    const detectedText = [...new Set(detected)];
    const detectedSummary = detectedText.length <= 3
      ? detectedText.join(', ')
      : `${detectedText.slice(0, 3).join(', ')} + ${detectedText.length - 3} more`;
    if (failures.length) {
      return this.finding(
        rule,
        'failed',
        failures.map(paragraph => paragraph.index),
        this.expected(rule),
        unavailable.length
          ? `${failures.length} failed; ${unavailable.length} unavailable`
          : detectedSummary,
        unavailable.length
          ? `${failures.length} paragraph(s) do not use ${expected} pt before and after spacing; ${unavailable.length} could not be evaluated.`
          : `${failures.length} paragraph(s) do not use ${expected} pt before and after spacing.`
      );
    }
    if (unavailable.length) {
      return this.finding(
        rule,
        'not_evaluated',
        unavailable.map(paragraph => paragraph.index),
        this.expected(rule),
        'Spacing unavailable',
        `${rule.name} could not be evaluated for ${unavailable.length} paragraph(s) because before/after spacing was unavailable.`,
        undefined,
        'technical'
      );
    }
    return this.finding(
      rule,
      'passed',
      [],
      this.expected(rule),
      detectedSummary || `${expected} pt before / ${expected} pt after`,
      `${targets.length} matching paragraph${targets.length === 1 ? '' : 's'} use ${expected} pt before and after spacing.`
    );
  }

  private static evaluateSingleLineSpacing(rule: RuleDefinition, snapshot: DocumentSnapshot): Finding {
    const eligible = snapshot.paragraphs.filter(p =>
      p.text.trim().length > 0
      && p.layoutSensitive !== true
      && p.blankParagraph?.hasDrawing !== true
    );
    const unreliable = eligible.filter(p => p.ooxmlAligned === false);
    if (unreliable.length) {
      return this.finding(
        rule,
        'not_evaluated',
        unreliable.map(p => p.index),
        this.expected(rule),
        'Paragraph mapping unavailable',
        'Line-spacing analysis was not scored because one or more meaningful paragraphs could not be mapped safely to OOXML.',
        undefined,
        'technical'
      );
    }
    const meaningful = eligible.filter(p => p.ooxmlAligned !== false);
    if (!meaningful.length) {
      return this.finding(rule, 'not_applicable', [], this.expected(rule), 'No meaningful paragraphs',
        'No manuscript paragraphs were detected, so line spacing does not apply.');
    }

    const unavailable = meaningful.filter(p =>
      p.lineSpacingMultiple === null && (!p.lineSpacingRule || !['exact', 'atleast'].includes(p.lineSpacingRule.toLowerCase()))
    );
    const offenders = meaningful.filter(p =>
      p.lineSpacingRule && ['exact', 'atleast'].includes(p.lineSpacingRule.toLowerCase())
      || (p.lineSpacingMultiple !== null && !approxEqual(p.lineSpacingMultiple, 1, 0.05))
    );
    if (offenders.length) {
      const details = offenders.map(p => this.lineSpacingDiagnostic(p)).join(' | ');
      return this.finding(
        rule,
        'failed',
        offenders.map(p => p.index),
        '1.0',
        unavailable.length
          ? `${offenders.length} failed; ${unavailable.length} unavailable. ${details}`
          : `${offenders.length} paragraph(s) use non-single spacing: ${details}`,
        unavailable.length
          ? `${offenders.length} paragraph(s) definitely fail; ${unavailable.length} could not be evaluated. ${details}`
          : `${offenders.length} paragraph(s) use non-single spacing. ${details}`
      );
    }

    if (unavailable.length) {
      return this.finding(
        rule,
        'not_evaluated',
        unavailable.map(p => p.index),
        '1.0',
        `${offenders.length} non-single paragraph(s); ${unavailable.length} unavailable`,
        'Word did not expose normalized line-spacing data for every meaningful paragraph.',
        undefined,
        'technical'
      );
    }

    return this.finding(
      rule,
      offenders.length ? 'failed' : 'passed',
      offenders.map(p => p.index),
      '1.0',
      offenders.length ? `${offenders.length} paragraph(s) use non-single spacing` : '1.0 throughout',
      offenders.length
        ? `${offenders.length} paragraph(s) use non-single spacing.`
        : `All ${meaningful.length} meaningful paragraphs use single spacing.`
    );
  }

  private static lineSpacingDiagnostic(paragraph: ParagraphSnapshot): string {
    const location = paragraph.classification === 'tableCaption'
      ? 'table caption'
      : paragraph.tableNestingLevel > 0
        ? 'table cell'
        : paragraph.classification === 'figureCaption'
          ? 'figure caption'
          : 'manuscript paragraph';
    const rule = paragraph.lineSpacingRule?.toLowerCase() || 'unknown';
    const value = paragraph.lineSpacingValue !== undefined && paragraph.lineSpacingValue !== null
      ? `${paragraph.lineSpacingValue} twips`
      : paragraph.lineSpacing !== null
        ? `${paragraph.lineSpacing} pt`
        : paragraph.lineSpacingMultiple !== null
          ? `${paragraph.lineSpacingMultiple}x`
          : 'unavailable';
    return `Paragraph ${paragraph.index + 1} (${location}): ${rule}, ${value}`;
  }

  private static evaluateRequiredSectionGroup(
    rule: RuleDefinition,
    snapshot: DocumentSnapshot,
    group: { label: string; families: readonly CanonicalSection[] }
  ): Finding {
    const paragraph = [...snapshot.paragraphs]
      .sort((a, b) => a.index - b.index)
      .find(p => isRecognizedFirstLevelSection(p)
        && p.sectionFamilies?.some(family => group.families.includes(family)));

    return this.finding(
      rule,
      paragraph ? 'passed' : 'failed',
      paragraph ? [paragraph.index] : [],
      this.expected(rule),
      paragraph ? paragraph.text.trim() : `No recognized ${group.label} section detected`,
      paragraph
        ? `${group.label} section detected as "${paragraph.text.trim()}".`
        : `No recognized ${group.label} section was detected.`
    );
  }

  private static evaluateCrossReferenceRule(rule: RuleDefinition, snapshot: DocumentSnapshot): Finding {
    const cross = snapshot.crossReferences;
    if (!cross) {
      return this.finding(rule, 'not_evaluated', [], this.expected(rule), 'Cross-reference data unavailable',
        'The document structure could not be analyzed for cross-references.', undefined, 'technical');
    }

    const definitionCount = rule.property.startsWith('crossref.reference')
      ? cross.referenceDefinitions.length
      : rule.property.startsWith('crossref.figure') || rule.property === 'structure.captionBelowFigure' || rule.property === 'structure.figurePriorCitation'
        ? cross.figureDefinitions.length
        : cross.tableDefinitions.length;
    if (definitionCount === 0) {
      return this.finding(rule, 'not_evaluated', [], this.expected(rule), 'Object numbering unavailable',
        'Matching content was detected, but no numbered definition could be resolved confidently from the OOXML.', undefined, 'technical');
    }

    if (rule.property.startsWith('crossref.reference') && cross.referenceCitationStyle === 'unknown') {
      return this.finding(rule, 'not_evaluated', [], this.expected(rule), 'No supported numeric citation syntax detected',
        'Reference entries were detected, but no square-bracket numeric citations were found. Citation coverage cannot be determined without a configured citation style.', undefined, 'technical');
    }

    let failedIndices: number[] = [];
    let unavailableIndices: number[] = [];
    let detected = 'No violations detected';
    const passMessage = `${rule.name} complies.`;
    let failMessage = `${rule.name} requires attention.`;

    switch (rule.property) {
      case 'crossref.referenceCoverage': {
        const missing = cross.referenceCoverageMissing;
        failedIndices = cross.referenceDefinitions.filter(d => missing.includes(d.number)).map(d => d.paragraphIndex);
        detected = missing.length ? `Uncited references: ${missing.join(', ')}` : 'All listed references are cited';
        failMessage = `${missing.length} listed reference(s) are not cited in the manuscript.`;
        break;
      }
      case 'crossref.referenceValid':
        failedIndices = [...new Set(cross.invalidReferenceCitations.map(c => c.paragraphIndex))];
        detected = cross.invalidReferenceCitations.length ? `${cross.invalidReferenceCitations.length} invalid reference citation occurrence(s)` : 'All reference citations are valid';
        failMessage = 'One or more citations do not correspond to a listed reference.';
        break;
      case 'crossref.referenceFirstUseSequence':
        failedIndices = [...new Set(cross.referenceSequenceViolations.map(c => c.paragraphIndex))];
        detected = failedIndices.length ? 'Reference first-use order is not sequential' : 'Reference first-use order is sequential';
        failMessage = 'Reference citations introduce new references out of sequence.';
        break;
      case 'crossref.figureCitationCoverage': {
        const missing = cross.figureCoverageMissing;
        failedIndices = cross.figureDefinitions.filter(d => missing.includes(d.number)).map(d => d.paragraphIndex);
        detected = missing.length ? `Uncited figures: ${missing.join(', ')}` : 'All figures are cited in text';
        failMessage = `${missing.length} figure(s) are not cited in the manuscript text.`;
        break;
      }
      case 'crossref.figureCitationSequence':
        failedIndices = [...new Set((cross.figureCitationSequenceViolations || []).map(citation => citation.paragraphIndex))];
        {
          const missing = this.missingSequenceNumbers(cross.figureDefinitions.map(definition => definition.number));
          detected = missing.length
            ? `Missing figure numbers: ${missing.join(', ')}; first citations follow the same gap`
            : failedIndices.length
              ? 'Figure first-citation order is not sequential'
              : 'Figure first-citation order is sequential';
          failMessage = missing.length
            ? `Figure captions/citations skip ${missing.join(', ')}. Confirm whether those figures were intentionally removed; renumbering is manual.`
            : 'Figure citations introduce first mentions out of numerical order.';
        }
        break;
      case 'crossref.figureCitationSpelling': {
        const violations = cross.figureCitations.filter(citation =>
          citation.expectedPrefix !== undefined
          && citation.originalText !== undefined
          && citation.originalText !== citation.expectedPrefix
        );
        const fieldBackedViolations = violations.filter(citation => citation.isFieldBacked === true);
        const safeViolations = violations.filter(citation => citation.isFieldBacked !== true);
        const fieldBackedIndices = [...new Set(fieldBackedViolations.map(citation => citation.paragraphIndex))];
        failedIndices = [...new Set(violations.map(citation => citation.paragraphIndex))];
        detected = violations.length
          ? `${violations.length} figure citation(s) use the wrong Figure/Fig. form${fieldBackedViolations.length ? `; ${fieldBackedViolations.length} field-backed citation(s) require manual review` : ''}`
          : cross.figureCitations.length ? 'Figure citation wording complies' : 'No figure citations detected';
        failMessage = 'One or more figure citations use an abbreviated Figure/Fig. form.';
        if (fieldBackedViolations.length && safeViolations.length === 0) {
          return this.finding(
            rule,
            'manual_review',
            fieldBackedIndices,
            this.expected(rule),
            detected,
            'One or more incorrect figure citations are field-backed Word content and must be corrected manually.',
            undefined,
            'manual',
            undefined,
            false
          );
        }
        if (violations.length) {
          const finding = this.finding(
            rule,
            'failed',
            failedIndices,
            this.expected(rule),
            detected,
            failMessage
          );
          finding.fixTextRanges = safeViolations.map(citation => ({
            paragraphIndex: citation.paragraphIndex,
            searchText: citation.originalCitationText!,
            originalPrefix: citation.originalText! as 'Figure' | 'Figures' | 'Fig.' | 'Figs.',
            replacementPrefix: citation.expectedPrefix! as 'Figure' | 'Figures' | 'Fig.' | 'Figs.',
            replacement: citation.correctedCitationText!,
            occurrenceIndex: citation.occurrenceIndex ?? 0
          }));
          return finding;
        }
        break;
      }
      case 'crossref.tableCitationCoverage': {
        const missing = cross.tableCoverageMissing;
        failedIndices = cross.tableDefinitions.filter(d => missing.includes(d.number)).map(d => d.paragraphIndex);
        detected = missing.length ? `Uncited tables: ${missing.join(', ')}` : 'All tables are cited in text';
        failMessage = `${missing.length} table(s) are not cited in the manuscript text.`;
        break;
      }
      case 'structure.captionBelowFigure': {
        const failures = cross.figurePlacement.filter(r => r.passed === false && r.confidence === 'high');
        const unknown = cross.figurePlacement.filter(r => r.passed === null || r.confidence !== 'high');
        failedIndices = [...new Set(failures.map(r => r.captionParagraphIndex))];
        unavailableIndices = [...new Set(unknown.map(r => r.captionParagraphIndex))];
        detected = failures.length ? `${failures.length} figure caption(s) precede their figure object` : 'Figure captions follow their detected figure objects';
        failMessage = 'One or more figure captions are not positioned below a confidently associated figure object.';
        break;
      }
      case 'structure.captionAboveTable': {
        const failures = cross.tablePlacement.filter(r => r.passed === false && r.confidence === 'high');
        const unknown = cross.tablePlacement.filter(r => r.passed === null || r.confidence !== 'high');
        failedIndices = [...new Set(failures.map(r => r.captionParagraphIndex))];
        unavailableIndices = [...new Set(unknown.map(r => r.captionParagraphIndex))];
        detected = failures.length ? `${failures.length} table caption(s) follow their table object` : 'Table captions precede their detected table objects';
        failMessage = 'One or more table captions are not positioned above a confidently associated table object.';
        break;
      }
      case 'structure.figurePriorCitation':
        failedIndices = [...new Set(cross.figurePriorCitationFailures.map(d => d.paragraphIndex))];
        unavailableIndices = [...new Set(cross.figurePriorCitationUnknown.map(d => d.paragraphIndex))];
        detected = failedIndices.length ? `${failedIndices.length} figure(s) lack a prior text citation` : 'Every confidently placed figure has a prior text citation';
        failMessage = 'One or more figures are placed before their first text citation.';
        break;
    }

    if (failedIndices.length) {
      return this.finding(rule, 'failed', failedIndices, this.expected(rule), detected,
        unavailableIndices.length ? `${failMessage} ${unavailableIndices.length} additional item(s) could not be evaluated.` : failMessage);
    }
    if (unavailableIndices.length) {
      return this.finding(rule, 'not_evaluated', unavailableIndices, this.expected(rule), 'Structure ambiguous',
        'The document contains object/caption relationships that could not be determined confidently.', undefined, 'technical');
    }
    return this.finding(rule, 'passed', [], this.expected(rule), detected, passMessage);
  }

  private static evaluateFigureSpaceAbove(rule: RuleDefinition, snapshot: DocumentSnapshot): Finding {
    const analysis = analyzeFigureSpaceAbove(snapshot);
    const actionable = analysis.issues.filter(issue => !issue.blockedByStructuralParagraph);
    if (actionable.length) {
      return this.finding(
        rule,
        'failed',
        actionable.map(issue => issue.captionParagraphIndex),
        this.expected(rule),
        `${actionable.length} inline figure(s) lack exactly one line space above`,
        `${actionable.length} confidently paired inline figure(s) require one blank paragraph before the figure object.`
      );
    }
    if (!analysis.applicableFigureCount) {
      return this.finding(
        rule,
        'not_evaluated',
        analysis.blockedCaptionIndices,
        this.expected(rule),
        'No safely auto-evaluable inline figure placement',
        'Only high-confidence single inline figures outside tables can be evaluated for automatic spacing above.',
        undefined,
        'technical'
      );
    }
    return this.finding(
      rule,
      'passed',
      [],
      this.expected(rule),
      'One line space present above each applicable figure',
      `${analysis.applicableFigureCount} applicable inline figure(s) comply.`
    );
  }

  private static evaluateTableCaptionTableGap(rule: RuleDefinition, snapshot: DocumentSnapshot): Finding {
    const analysis = analyzeTableCaptionTableGaps(snapshot);
    const actionable = analysis.issues.filter(issue => !issue.blockedByStructuralParagraph);
    const blocked = analysis.issues.filter(issue => issue.blockedByStructuralParagraph);
    const extraBlankCount = actionable.reduce(
      (count, issue) => count + issue.blankParagraphIndices.length,
      0
    );

    if (actionable.length) {
      return this.finding(
        rule,
        'failed',
        actionable.map(issue => issue.captionParagraphIndex),
        this.expected(rule),
        `${extraBlankCount} extra blank paragraph${extraBlankCount === 1 ? '' : 's'} between table caption and table`,
        `${extraBlankCount} safely removable blank paragraph${extraBlankCount === 1 ? '' : 's'} separate a table caption from its associated table.`
      );
    }

    if (blocked.length) {
      return this.finding(
        rule,
        'not_evaluated',
        blocked.map(issue => issue.captionParagraphIndex),
        this.expected(rule),
        'Structural paragraph prevents safe table-caption gap cleanup',
        'A structural break or protected paragraph occurs between a table caption and its table; no automatic deletion was made.',
        undefined,
        'technical'
      );
    }

    if (!analysis.applicableCaptionCount) {
      return this.finding(
        rule,
        'not_applicable',
        [],
        this.expected(rule),
        'No high-confidence external table-caption/table relationship detected',
        'Embedded captions, ambiguous placements, and captions separated by content are excluded from this structural gap rule.'
      );
    }

    return this.finding(
      rule,
      'passed',
      [],
      this.expected(rule),
      'Table captions directly precede their associated tables',
      `${analysis.applicableCaptionCount} applicable external table caption${analysis.applicableCaptionCount === 1 ? '' : 's'} comply.`
    );
  }

  private static evaluateSectionBodyTypography(
    rule: RuleDefinition,
    snapshot: DocumentSnapshot,
    headingText: string
  ): Finding {
    const heading = snapshot.paragraphs.find(p =>
      p.classification === 'heading1' && p.text.trim().toUpperCase() === headingText
    );
    if (!heading) {
      return this.finding(rule, 'not_applicable', [], this.expected(rule), `${headingText} section not detected`,
        `The ${headingText} section is not present, so this rule does not apply.`);
    }

    const sectionBody = snapshot.paragraphs.filter(p =>
      p.index > heading.index
      && p.classification === 'body'
      && isReliableLayoutParagraph(p)
      && !snapshot.paragraphs.some(h => h.classification === 'heading1' && h.index > heading.index && h.index < p.index)
      && (snapshot.referencesHeadingIndex === null || p.index < snapshot.referencesHeadingIndex)
    );

    const unreliableSectionBody = snapshot.paragraphs.filter(p =>
      p.index > heading.index
      && p.classification === 'body'
      && p.ooxmlAligned === false
      && !snapshot.paragraphs.some(h => h.classification === 'heading1' && h.index > heading.index && h.index < p.index)
      && (snapshot.referencesHeadingIndex === null || p.index < snapshot.referencesHeadingIndex)
    );
    if (unreliableSectionBody.length) {
      return this.finding(
        rule,
        'not_evaluated',
        unreliableSectionBody.map(p => p.index),
        this.expected(rule),
        'Paragraph mapping unavailable',
        `${headingText} body typography was not scored because one or more section paragraphs could not be mapped safely to OOXML.`,
        undefined,
        'technical'
      );
    }

    if (!sectionBody.length) {
      return this.finding(rule, 'not_applicable', [], this.expected(rule), 'No section body paragraphs detected',
        `${headingText} was detected but no body paragraph was found before the next first-level heading.`);
    }

    const failed = sectionBody.filter(p =>
      (p.fontName || '').toLowerCase() !== 'times new roman' || !approxEqual(p.fontSize, 10, 0.75)
    );
    const pass = failed.length === 0;
    return this.finding(
      rule,
      pass ? 'passed' : 'failed',
      failed.map(p => p.index),
      'Times New Roman, 10 pt',
      pass ? 'Times New Roman, 10 pt' : `${failed.length} section paragraph(s) use different typography`,
      pass ? `${headingText} body typography complies.` : `${failed.length} ${headingText} body paragraph(s) require attention.`
    );
  }

  private static booleanDocumentFinding(
    rule: RuleDefinition,
    pass: boolean,
    indices: number[],
    detected: string,
    passMessage: string,
    failMessage: string
  ): Finding {
    return this.finding(rule, pass ? 'passed' : 'failed', indices, this.expected(rule), detected, pass ? passMessage : failMessage);
  }

  private static evaluateTextOrStructure(rule: RuleDefinition, targets: ParagraphSnapshot[], snapshot: DocumentSnapshot): Finding {
    let failed: ParagraphSnapshot[] = [];
    let detected = '';

    switch (rule.property) {
      case 'text.notAllCaps':
        failed = targets.filter(p => isAllCaps(p.text) || p.allCaps === true);
        detected = failed.length ? 'ALL CAPS detected' : 'Not all caps';
        break;
      case 'text.allCaps':
        failed = targets.filter(p => !isAllCaps(p.text) && p.allCaps !== true);
        detected = failed.length ? 'One or more headings are not all caps' : 'All caps';
        break;
      case 'text.titleCase':
        failed = targets.filter(p => !looksMajorTitleCase(this.cleanGuidanceSuffix(p.text)));
        detected = failed.length ? 'Major-word capitalization mismatch detected' : 'Major words use initial-letter capitalization';
        break;
      case 'text.minorWordsLowercase':
        failed = targets.filter(p => !this.minorWordsAreLowercase(this.cleanGuidanceSuffix(p.text)));
        detected = failed.length ? 'Capitalized preposition/article detected' : 'Minor words lowercase';
        break;
      case 'text.abstractPrefix':
        failed = targets.filter(p => !/^\s*Abstract\b/i.test(p.text));
        detected = failed.length ? 'Missing “Abstract” prefix' : '“Abstract” prefix detected';
        break;
      case 'text.leadingWhitespace':
        failed = targets.filter(p => /^[ \t]/.test(p.text));
        detected = failed.length ? `${failed.length} caption(s) contain leading spaces or tabs` : 'No leading spaces or tabs';
        break;
      case 'text.abstractPrefixPeriod':
        failed = targets.filter(p => !/^\s*Abstract\s*\./i.test(p.text));
        detected = failed.length ? 'Missing period after “Abstract”' : '“Abstract.” prefix detected';
        break;
      case 'run.abstractPrefixBold':
        failed = targets.filter(p => p.abstractPrefixBold !== true);
        detected = failed.length ? '“Abstract.” prefix is not fully bold' : '“Abstract.” prefix bold';
        if (failed.length) {
          const mappingReliable = snapshot.layout.paragraphMappingAligned !== false;
          const safe = failed.filter(paragraph => isSafeAbstractPrefixBoldCandidate(paragraph, mappingReliable));
          const manual = failed.filter(paragraph => !safe.includes(paragraph));
          if (!safe.length) {
            return this.finding(
              rule,
              'manual_review',
              failed.map(paragraph => paragraph.index),
              this.expected(rule),
              manual.some(paragraph => hasFieldBackedAbstractPrefix(paragraph))
                ? 'Abstract prefix is field-backed or cannot be safely mapped'
                : 'Exact “Abstract.” prefix is missing or cannot be safely mapped',
              'The Abstract prefix must be corrected manually because the exact prefix, field mapping, or paragraph mapping could not be verified safely.',
              undefined,
              'manual',
              undefined,
              false
            );
          }
          const finding = this.finding(
            rule,
            'failed',
            failed.map(paragraph => paragraph.index),
            this.expected(rule),
            manual.length
              ? `${safe.length} safe prefix${safe.length === 1 ? '' : 'es'}; ${manual.length} require manual review`
              : detected,
            manual.length
              ? `${safe.length} Abstract prefix${safe.length === 1 ? '' : 'es'} can be bolded safely; ${manual.length} require manual review.`
              : 'The exact Abstract. prefix is not bold and can be corrected without changing the abstract body.',
            safe.map(paragraph => paragraph.index),
            undefined,
            undefined,
            true
          );
          finding.fixParagraphTexts = Object.fromEntries(
            safe.map(paragraph => [paragraph.index, paragraph.text])
          );
          return finding;
        }
        break;
      case 'run.abstractBodyBold':
        // A mixed run range is a confirmed violation, even though Word reports
        // its aggregate bold value as null. Only null/null means unavailable.
        failed = targets.filter(p => p.abstractBodyBold === true || p.abstractBodyMixed === true);
        const unavailableAbstractBody = targets.filter(
          p => p.abstractBodyBold === null && p.abstractBodyMixed !== true
        );
        if (failed.length) {
          return this.finding(
            rule,
            'failed',
            failed.map(p => p.index),
            this.expected(rule),
            unavailableAbstractBody.length
              ? `${failed.length} failed; ${unavailableAbstractBody.length} unavailable`
              : 'Abstract body text is bold or mixed formatting',
            unavailableAbstractBody.length
              ? `${failed.length} Abstract paragraph(s) definitely fail; ${unavailableAbstractBody.length} could not be evaluated.`
              : 'Abstract body contains bold or mixed formatting.'
          );
        }
        if (unavailableAbstractBody.length) {
          return this.finding(
            rule,
            'not_evaluated',
            unavailableAbstractBody.map(p => p.index),
            this.expected(rule),
            'Formatting unavailable',
            'Abstract body run formatting could not be read from Word.',
            undefined,
            'technical'
          );
        }
        detected = 'Abstract body text is normal';
        break;
      case 'text.figurePrefix':
        failed = targets.filter(
          p => !/^\s*FIGURE\s+\d+[A-Z]?\s*\./.test(p.text) && p.figurePrefixAllCaps !== true
        );
        detected = failed.length ? 'Invalid figure-label prefix' : 'FIGURE n. detected';
        if (failed.length) {
          const malformed = failed.filter(hasUnparseableFigureCaptionLabel);
          const fieldBacked = failed.filter(hasFieldBackedFigureCaptionLabel);
          const manualOnly = [...new Set([...malformed, ...fieldBacked])];
          const safe = failed.filter(p => !manualOnly.includes(p));
          if (manualOnly.length && safe.length === 0) {
            return this.finding(
              rule,
              'manual_review',
              manualOnly.map(p => p.index),
              this.expected(rule),
              malformed.length
                ? 'Figure caption label is malformed and cannot be parsed safely'
                : 'Figure caption label is field-backed Word content',
              malformed.length
                ? 'The caption already begins with FIGURE/Fig., but its label format is ambiguous and must be corrected manually.'
                : 'The figure caption label contains a Word field and must be corrected manually.',
              undefined,
              'manual',
              undefined,
              false
            );
          }
          if (manualOnly.length) {
            const reasons = [
              malformed.length ? `${malformed.length} malformed label${malformed.length === 1 ? '' : 's'}` : '',
              fieldBacked.length ? `${fieldBacked.length} field-backed label${fieldBacked.length === 1 ? '' : 's'}` : ''
            ].filter(Boolean).join('; ');
            detected = `${failed.length} invalid figure-label prefix${failed.length === 1 ? '' : 'es'}; ${reasons} require manual review`;
          }
        }
        break;
      case 'run.figurePrefixBold':
        failed = targets.filter(p => p.figurePrefixBold !== true);
        detected = failed.length ? 'One or more FIGURE n. labels are not bold' : 'FIGURE n. labels bold';
        break;
      case 'text.tablePrefix':
        failed = targets.filter(
          p => !/^\s*TABLE\s+\d+[A-Z]?\s*\./i.test(p.text)
        );
        detected = failed.length ? 'Invalid table-label prefix' : 'TABLE n. detected';
        if (failed.length) {
          const safeRoman = failed.filter(isSafeRomanTableCaptionFix);
          const manual = failed.filter(p => !safeRoman.includes(p));
          if (safeRoman.length) {
            return this.finding(
              rule,
              'failed',
              failed.map(p => p.index),
              this.expected(rule),
              manual.length
                ? `${safeRoman.length} Roman table label(s) can be normalized; ${manual.length} require manual review`
                : 'Roman table label can be normalized safely',
              manual.length
                ? `${safeRoman.length} parseable Roman table label(s) are safe to convert to Arabic numbering; other malformed or field-backed labels require manual review.`
                : 'A parseable Roman table label can be converted to its Arabic equivalent without changing the caption body.',
              safeRoman.map(p => p.index),
              undefined,
              undefined,
              true
            );
          }
        }
        break;
      case 'run.tablePrefixBold':
        failed = targets.filter(p => p.tablePrefixBold !== true);
        detected = failed.length ? 'One or more TABLE n. labels are not bold' : 'TABLE n. labels bold';
        if (failed.length) {
          const safe = failed.filter(paragraph => isSafeTablePrefixBoldCandidate(paragraph, snapshot));
          const manual = failed.filter(paragraph => !safe.includes(paragraph));
          if (!safe.length) {
            return this.finding(
              rule,
              'manual_review',
              failed.map(paragraph => paragraph.index),
              this.expected(rule),
              manual.some(paragraph => hasFieldBackedTableCaptionPrefix(paragraph))
                ? 'TABLE label is field-backed or cannot be safely mapped'
                : 'Exact TABLE n. prefix is missing or cannot be safely mapped',
              'The TABLE prefix must be corrected manually because its exact range, field mapping, or paragraph mapping could not be verified safely.',
              undefined,
              'manual',
              undefined,
              false
            );
          }
          const finding = this.finding(
            rule,
            'failed',
            failed.map(paragraph => paragraph.index),
            this.expected(rule),
            manual.length
              ? `${safe.length} safe TABLE prefix${safe.length === 1 ? '' : 'es'}; ${manual.length} require manual review`
              : detected,
            manual.length
              ? `${safe.length} TABLE prefix${safe.length === 1 ? '' : 'es'} can be bolded safely; ${manual.length} require manual review.`
              : 'The exact TABLE n. prefixes are not bold and can be corrected without changing caption body text.',
            safe.map(paragraph => paragraph.index),
            undefined,
            undefined,
            true
          );
          finding.fixParagraphTexts = Object.fromEntries(
            safe.map(paragraph => [paragraph.index, paragraph.text])
          );
          return finding;
        }
        break;
      case 'text.tablePrefixCaps':
        failed = targets.filter(
          p => p.tablePrefixAllCaps !== true
        );
        detected = failed.length ? 'TABLE label capitalization mismatch' : 'TABLE label uppercase';
        break;
      case 'text.referenceNumbering': {
        failed = this.sequenceFailures(targets, p => {
          const listMatch = (p.listString || '').trim().match(/^(\d+)\.$/);
          const metadata = p.referenceNumbering;
          const automatic = metadata?.kind === 'automatic'
            || (metadata === undefined && p.isListItem);
          const arabic = metadata?.format === undefined
            || metadata?.format === null
            || metadata.format === 'decimal';
          const period = metadata?.suffix === undefined
            || metadata?.suffix === null
            || metadata.suffix === '.'
            || Boolean(listMatch);
          const numberAlignment = metadata?.alignment === undefined
            || metadata?.alignment === null
            || normalizeAlignment(metadata.alignment) === 'left';
          const startsAtOne = metadata?.startAt === undefined
            || metadata?.startAt === null
            || metadata.startAt === 1;
          if (!automatic || !arabic || !period || !numberAlignment || !startsAtOne || metadata?.kind === 'mixed') return null;
          return listMatch ? Number(listMatch[1]) : null;
        });
        detected = failed.length ? 'Missing or non-sequential automatic Arabic “1., 2., 3.” numbering' : 'Automatic sequential Arabic numerals + period';
        break;
      }
      case 'reference.style':
        {
        const wrongStyle = targets.filter(p => referenceStyleState(p) === 'wrong');
        if (wrongStyle.length) {
          return this.finding(
            rule,
            'failed',
            wrongStyle.map(p => p.index),
            this.expected(rule),
            'One or more bibliography paragraphs use a style other than Reference',
            'Each bibliography paragraph must use the Reference paragraph style.'
          );
        }
        if (targets.some(p => referenceStyleState(p) === 'unknown' || !p.referenceStyle)) {
          return this.finding(
            rule,
            'not_evaluated',
            targets.filter(p => referenceStyleState(p) !== 'reference' || !p.referenceStyle).map(p => p.index),
            this.expected(rule),
            'Reference style metadata unavailable',
            'The Reference style could not be read from styles.xml, so its type, inheritance, and next-paragraph settings were not scored safely.',
            undefined,
            'technical'
          );
        }
        failed = targets.filter(p => {
          const style = p.referenceStyle;
          const name = style?.name || p.style;
          const type = style?.type;
          const basedOn = style?.basedOn;
          // Word treats an omitted w:next on a paragraph style as “continue
          // with this style”.  The AIP Reference style intentionally omits
          // w:next, so absence is not evidence of a wrong style.
          const next = style?.nextParagraphStyle || style?.name || p.style;
          return !/^reference$/i.test(name || '')
            || !/^paragraph$/i.test(type || '')
            || !/^paragraph$/i.test(basedOn || '')
            || !/^reference$/i.test(next || '');
        });
        detected = failed.length ? 'Reference paragraph style is missing or has incorrect metadata' : 'Reference paragraph style';
        break;
        }
      case 'reference.style.fontName':
      case 'reference.style.fontSize':
      case 'reference.style.alignment':
      case 'reference.style.leftIndent':
      case 'reference.style.hangingIndent':
      case 'reference.list.level':
      case 'reference.list.numberPosition':
      case 'reference.list.textPosition':
        {
        const wrongStyle = targets.filter(p => referenceStyleState(p) === 'wrong');
        if (wrongStyle.length && rule.property.startsWith('reference.style.')) {
          return this.finding(
            rule,
            'failed',
            wrongStyle.map(p => p.index),
            this.expected(rule),
            'One or more bibliography paragraphs use a style other than Reference',
            'The Reference style formatting rules cannot pass while the paragraph uses Normal or another non-Reference style.'
          );
        }
        if (targets.some(p =>
          (rule.property.startsWith('reference.style.')
            && (referenceStyleState(p) !== 'reference' || !p.referenceStyle))
          || (rule.property.startsWith('reference.list.') && !p.referenceNumbering && !p.isListItem)
        )) {
          return this.finding(
            rule,
            'not_evaluated',
            targets.filter(p =>
              (rule.property.startsWith('reference.style.') && !p.referenceStyle)
              || (rule.property.startsWith('reference.list.') && !p.referenceNumbering && !p.isListItem)
            ).map(p => p.index),
            this.expected(rule),
            'Reference formatting metadata unavailable',
            'The required Reference style or list-level metadata was not available from the live document and OOXML.',
            undefined,
            'technical'
          );
        }
        failed = targets.filter(p => {
          const styleValue = this.paragraphProperty(p, rule.property);
          const isIndentationRule = rule.property === 'reference.style.leftIndent'
            || rule.property === 'reference.style.hangingIndent';
          const effectiveValue = rule.property.startsWith('reference.style.') && !isIndentationRule
            ? this.referenceEffectiveProperty(p, rule.property)
            : styleValue;
          const runValues = this.referenceRunPropertyValues(p, rule.property);
          const runMismatch = runValues !== null
            && runValues.some(value => value === null
              || value === undefined
              || !this.matches(value, rule.expected, rule.property));
          const directIndentOverride = isIndentationRule
            && (p.formattingSources?.leftIndent === 'direct'
              || p.formattingSources?.firstLineIndent === 'direct');
          return styleValue === null
            || styleValue === undefined
            || !this.matches(styleValue, rule.expected, rule.property)
            || effectiveValue === null
            || effectiveValue === undefined
            || !this.matches(effectiveValue, rule.expected, rule.property)
            || runMismatch
            || directIndentOverride;
        });
        detected = failed.length ? `${failed.length} reference formatting layer value(s) do not match` : `${rule.name} complies`;
        break;
        }
      case 'text.figureNumberSequence': {
        const sequence = this.figureNumberSequence(targets);
        failed = sequence.failures;
        detected = sequence.missing.length
          ? `Missing figure numbers: ${sequence.missing.join(', ')}${sequence.mismatches.length ? `; ${sequence.mismatches.join('; ')}` : ''}`
          : sequence.duplicates.length
            ? `Duplicate figure numbers: ${sequence.duplicates.join(', ')}`
            : failed.length
              ? 'Figure captions appear out of numerical order'
              : 'Sequential FIGURE 1., FIGURE 2., … numbering';
        break;
      }
      case 'text.tableNumberSequence':
        failed = this.sequenceFailures(targets, p => parseCaptionNumber(p.text, 'table'));
        detected = failed.length ? 'Non-sequential table numbering detected' : 'Sequential TABLE 1., TABLE 2., … numbering';
        break;
      case 'text.equationNumberPresent':
      case 'text.equationNumberParentheses':
        failed = targets.filter(p => !/\(\s*\d+\s*\)\s*$/.test(p.text));
        detected = failed.length ? 'Equation number in parentheses missing' : 'Parenthesized equation number detected';
        break;
      case 'text.equationNumberSequence':
        failed = this.sequenceFailures(targets, p => Number(p.text.match(/\(\s*(\d+)\s*\)\s*$/)?.[1] || NaN));
        detected = failed.length ? 'Equation numbering is missing or non-sequential' : 'Sequential equation numbers detected';
        break;
      case 'structure.notAfterReferences':
        if (snapshot.referencesHeadingIndex === null) {
          return this.finding(rule, 'not_applicable', [], this.expected(rule), 'References heading not detected',
            'This structure rule depends on a REFERENCES section; because none was detected, it was not scored.');
        }
        failed = targets.filter(p => p.index > snapshot.referencesHeadingIndex!);
        detected = failed.length ? `${failed.length} caption(s) after REFERENCES` : 'No figure captions after REFERENCES';
        break;
      case 'structure.blankLineAfterCaptionBeforeBody': {
        const analysis = analyzeFigureCaptionBodyGaps(snapshot.paragraphs);
        const issues = analysis.issues;
        const actionableIssues = issues.filter(issue => !issue.blockedByStructuralParagraph);
        const blockedIssues = issues.filter(issue => issue.blockedByStructuralParagraph);
        const issueIndices = new Set(actionableIssues.map(issue => issue.captionParagraphIndex));
        failed = targets.filter(paragraph => issueIndices.has(paragraph.index));
        const missing = actionableIssues.filter(issue => issue.blankParagraphIndices.length === 0).length;
        const extra = actionableIssues.reduce(
          (count, issue) => count + Math.max(0, issue.blankParagraphIndices.length - 1),
          0
        );
        detected = actionableIssues.length === 0
          ? blockedIssues.length
            ? 'Structural break prevents safe caption-gap evaluation'
            : 'Exactly one safe blank paragraph separates each applicable figure caption from body text'
          :
            `${missing} caption gap(s) missing; ${extra} extra blank paragraph(s) detected`;
        if (!failed.length && analysis.blockedCaptionIndices.length) {
          return this.finding(
            rule,
            'not_evaluated',
            blockedIssues.map(issue => issue.captionParagraphIndex),
            this.expected(rule),
            detected,
            'A page break, section break, manual break, or other structural paragraph prevents safe automatic caption-gap normalization.',
            undefined,
            'technical'
          );
        }
        if (!failed.length && analysis.applicableCaptionCount === 0) {
          return this.finding(
            rule,
            'not_applicable',
            [],
            this.expected(rule),
            'No figure caption is followed by body text',
            'No applicable figure-caption-to-body relationship was detected; headings, tables, other captions, and terminal captions are excluded.'
          );
        }
        break;
      }
      case 'structure.authorAffiliationSuperscripts':
        return this.evaluateAssociationSuperscripts(rule, targets, 'affiliation');
      case 'structure.authorEmailSuperscripts':
        return this.evaluateAssociationSuperscripts(rule, targets, 'email');
      case 'structure.equationTabs':
        return this.evaluateEquationTabs(rule, targets);
      case 'structure.tableTopRule':
      case 'structure.tableBottomRule':
      case 'structure.tableHeaderRule':
      case 'structure.tableNoVerticalRules':
      case 'structure.tableNoExtraHorizontalRules':
        return this.evaluateTableRule(rule, targets, snapshot);
      default:
        return this.finding(rule, 'not_evaluated', [], this.expected(rule), 'Not evaluated',
          'This deterministic text/structure rule is not yet implemented.', undefined, 'technical');
    }

    const pass = failed.length === 0;
    const fixParagraphIndices = rule.property === 'text.figurePrefix' || rule.property === 'text.referenceNumbering'
      ? targets.map(p => p.index)
      : failed.map(p => p.index);
    return this.finding(
      rule,
      pass ? 'passed' : 'failed',
      failed.map(p => p.index),
      this.expected(rule),
      detected,
      pass ? `${rule.name} complies.` : `${failed.length} matching paragraph${failed.length === 1 ? '' : 's'} require attention.`,
      fixParagraphIndices
    );
  }

  private static evaluateAssociationSuperscripts(
    rule: RuleDefinition,
    targets: ParagraphSnapshot[],
    kind: 'affiliation' | 'email'
  ): Finding {
    const failures: ParagraphSnapshot[] = [];
    const unavailable: ParagraphSnapshot[] = [];
    let missingMarker = false;
    for (const paragraph of targets) {
      const logicalText = this.normalizeLogicalLineBreaks(paragraph.text);
      const lines = this.logicalLines(logicalText);

      // Affiliations are defined by the paragraph-leading numeric marker.
      // Email paragraphs can contain an affiliation line followed by several
      // manual-break email lines, so email markers are evaluated per logical
      // line instead of only at the Word paragraph start.
      const candidates = kind === 'affiliation'
        ? (() => {
            const match = logicalText.match(/^(\s*)(\d+(?:\s*[,;]\s*\d+)*)/);
            return match
              ? [{ start: match[1].length, marker: match[2] }]
              : [];
          })()
        : lines
            .filter(line => /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i.test(line.text) || /^\s*[A-Za-z]\)/.test(line.text))
            .map(line => {
              const match = line.text.match(/^(\s*)([A-Za-z]\))/);
              return match
                ? { start: line.start + match[1].length, marker: match[2] }
                : { start: line.start, marker: null };
            });

      if (!candidates.length) {
        failures.push(paragraph);
        missingMarker = true;
        continue;
      }
      if (!paragraph.effectiveRuns) {
        unavailable.push(paragraph);
        continue;
      }

      let paragraphFailed = false;
      let paragraphUnavailable = false;
      for (const candidate of candidates) {
        if (!candidate.marker) {
          paragraphFailed = true;
          missingMarker = true;
          continue;
        }
        const markerEnd = candidate.start + candidate.marker.length;
        let offset = 0;
        let covered = 0;
        let valid = true;
        for (const run of paragraph.effectiveRuns) {
          const runText = this.normalizeLogicalLineBreaks(run.text);
          const start = offset;
          const end = offset + runText.length;
          const overlap = Math.max(0, Math.min(end, markerEnd) - Math.max(start, candidate.start));
          if (overlap > 0) {
            covered += overlap;
            if (run.verticalAlign !== 'superscript') valid = false;
          }
          offset = end;
        }
        if (covered < candidate.marker.length) paragraphUnavailable = true;
        else if (!valid) paragraphFailed = true;
      }
      if (paragraphFailed) failures.push(paragraph);
      else if (paragraphUnavailable) unavailable.push(paragraph);
    }
    if (failures.length) {
      return this.finding(rule, 'failed', failures.map(p => p.index), this.expected(rule),
        unavailable.length ? `${failures.length} failed; ${unavailable.length} unavailable` : 'One or more association markers are not superscript',
        unavailable.length ? `${failures.length} marker(s) fail; ${unavailable.length} could not be evaluated.` : `${failures.length} author association marker(s) require superscript formatting.`,
        undefined,
        undefined,
        undefined,
        missingMarker ? false : undefined);
    }
    if (unavailable.length) {
      return this.finding(rule, 'not_evaluated', unavailable.map(p => p.index), this.expected(rule), 'Marker formatting unavailable',
        'The association marker was detected, but run-level formatting was unavailable.', undefined, 'technical');
    }
    return this.finding(rule, 'passed', [], this.expected(rule), 'All association markers are superscript', `${rule.name} complies.`);
  }

  private static normalizeLogicalLineBreaks(text: string): string {
    return text.replace(/\r\n/g, '\n').replace(/[\r\v\f]/g, '\n');
  }

  private static logicalLines(text: string): Array<{ text: string; start: number }> {
    const lines: Array<{ text: string; start: number }> = [];
    let start = 0;
    for (const line of text.split('\n')) {
      lines.push({ text: line, start });
      start += line.length + 1;
    }
    return lines;
  }

  private static evaluateEquationTabs(rule: RuleDefinition, targets: ParagraphSnapshot[]): Finding {
    const failures: ParagraphSnapshot[] = [];
    const unavailable: ParagraphSnapshot[] = [];
    for (const paragraph of targets) {
      if (!paragraph.tabStops) {
        unavailable.push(paragraph);
        continue;
      }
      const hasCenter = paragraph.tabStops.some(tab => tab.alignment === 'center' && Math.abs(tab.positionPt - 216) <= 0.75);
      const hasRight = paragraph.tabStops.some(tab => tab.alignment === 'right' && Math.abs(tab.positionPt - (9242 / 20)) <= 0.75);
      if (!hasCenter || !hasRight) failures.push(paragraph);
    }
    if (failures.length) {
      return this.finding(rule, 'failed', failures.map(p => p.index), this.expected(rule),
        unavailable.length ? `${failures.length} failed; ${unavailable.length} unavailable` : 'Center and right tab stops missing',
        unavailable.length ? `${failures.length} equation(s) fail; ${unavailable.length} could not be evaluated.` : `${failures.length} equation paragraph(s) do not have both center and right tab stops.`);
    }
    if (unavailable.length) return this.finding(rule, 'not_evaluated', unavailable.map(p => p.index), this.expected(rule), 'Tab stops unavailable', 'Equation tab stops could not be read from OOXML.', undefined, 'technical');
    return this.finding(rule, 'passed', [], this.expected(rule), 'Center and right tab stops detected', `${rule.name} complies.`);
  }

  private static evaluateTableRule(rule: RuleDefinition, targets: ParagraphSnapshot[], snapshot: DocumentSnapshot): Finding {
    const property = rule.property;
    const failures: ParagraphSnapshot[] = [];
    const unavailable: ParagraphSnapshot[] = [];
    const detected: string[] = [];
    const tableIndices: number[] = [];
    const uniqueTables = new Map<number, { caption: ParagraphSnapshot; table: NonNullable<DocumentSnapshot['layout']['documentBlocks']>[number] }>();
    for (const caption of targets) {
      const relation = snapshot.crossReferences?.tablePlacement.find(item => item.captionParagraphIndex === caption.index);
      const table = relation && relation.confidence === 'high' && relation.objectBlockIndex !== null
        ? snapshot.layout.documentBlocks?.find(block => block.blockIndex === relation.objectBlockIndex)
        : undefined;
      if (typeof table?.tableIndex === 'number') {
        if (!uniqueTables.has(table.tableIndex)) uniqueTables.set(table.tableIndex, { caption, table });
        continue;
      }
      unavailable.push(caption);
    }
    for (const { caption, table } of uniqueTables.values()) {
      tableIndices.push(table.tableIndex as number);
      const value = property === 'structure.tableTopRule'
        ? table?.tableTopBorder
        : property === 'structure.tableBottomRule'
          ? table?.tableBottomBorder
          : property === 'structure.tableHeaderRule'
            ? table?.tableHeaderRule
            : property === 'structure.tableNoVerticalRules'
              ? table?.tableVerticalRules === undefined ? undefined : !table.tableVerticalRules
              : table?.tableExtraHorizontalRules === undefined ? undefined : !table.tableExtraHorizontalRules;
      if (value === true) {
        detected.push('present');
      } else if (value === false) {
        failures.push(caption);
        detected.push(property === 'structure.tableNoVerticalRules'
          ? 'vertical rules detected'
          : property === 'structure.tableNoExtraHorizontalRules'
            ? 'extra horizontal rules detected'
            : 'missing');
      } else {
        unavailable.push(caption);
      }
    }
    if (failures.length) {
      const failDetected = property === 'structure.tableHeaderRule'
        ? `${failures.length} table(s) missing the required header separator`
        : property === 'structure.tableNoExtraHorizontalRules'
          ? `${failures.length} table(s) contain extra horizontal body rules`
          : property === 'structure.tableNoVerticalRules'
            ? `${failures.length} table(s) contain vertical rules`
            : property === 'structure.tableTopRule'
              ? `${failures.length} table(s) missing the required top rule`
              : `${failures.length} table(s) missing the required bottom rule`;
      const failMessage = property === 'structure.tableHeaderRule'
        ? `${failures.length} table caption(s) are associated with tables missing the required header separator.`
        : property === 'structure.tableNoExtraHorizontalRules'
          ? `${failures.length} table caption(s) are associated with tables containing extra horizontal body rules.`
          : property === 'structure.tableNoVerticalRules'
            ? `${failures.length} table caption(s) are associated with tables containing vertical rules.`
            : `${failures.length} table caption(s) are associated with tables missing the required ${property === 'structure.tableTopRule' ? 'top' : 'bottom'} border rule.`;
      const failedAssociations = failures
        .map(caption => snapshot.crossReferences?.tablePlacement.find(item => item.captionParagraphIndex === caption.index)?.objectBlockIndex)
        .map(blockIndex => snapshot.layout.documentBlocks?.find(block => block.blockIndex === blockIndex))
        .filter((block): block is NonNullable<typeof block> => typeof block?.tableIndex === 'number');
      const failedTableIndices = failedAssociations.map(block => block.tableIndex as number);
      const failedHeaderRowCounts = failedAssociations.map(block => Math.max(1, block.tableHeaderRowCount || 1));
      const failedCaptionRows = failedAssociations.map(block => block.tableCaptionRowIndex ?? null);
      const failedHeaderStarts = failedAssociations.map(block => block.tableHeaderStartRowIndex ?? null);
      const failedHeaderEnds = failedAssociations.map(block => block.tableHeaderEndRowIndex ?? null);
      return this.finding(
        rule,
        'failed',
        failures.map(p => p.index),
        this.expected(rule),
        failDetected,
        failMessage,
        undefined,
        undefined,
        failedTableIndices,
        undefined,
        failedHeaderRowCounts,
        failedCaptionRows,
        failedHeaderStarts,
        failedHeaderEnds
      );
    }
    if (unavailable.length) return this.finding(rule, 'not_evaluated', unavailable.map(p => p.index), this.expected(rule), 'Table border data unavailable', 'The table could not be associated or its OOXML border data was unavailable.', undefined, 'technical');
    return this.finding(rule, 'passed', [], this.expected(rule), 'Rule present for all tables', `${rule.name} complies.`, undefined, undefined, tableIndices);
  }

  private static sequenceFailures(
    targets: ParagraphSnapshot[],
    getNumber: (p: ParagraphSnapshot) => number | null
  ): ParagraphSnapshot[] {
    const failed: ParagraphSnapshot[] = [];
    targets.forEach((p, index) => {
      const actual = getNumber(p);
      if (!Number.isFinite(actual) || actual !== index + 1) failed.push(p);
    });
    return failed;
  }

  private static minorWordsAreLowercase(text: string): boolean {
    const words = text.match(/[A-Za-z]+(?:['’][A-Za-z]+)?/g) || [];
    if (words.length < 3) return true;
    for (let i = 1; i < words.length - 1; i++) {
      const word = words[i];
      if (SMALL_HEADING_WORDS.has(word.toLowerCase()) && word !== word.toLowerCase()) return false;
    }
    return true;
  }

  private static cleanGuidanceSuffix(text: string): string {
    return text.replace(/\s*\((First|Second|Third) Level Heading\)\s*$/i, '').trim();
  }

  private static normalizeTemplateText(text: string): string {
    return text
      .normalize('NFKC')
      .replace(/[“”„‟]/g, '"')
      .replace(/[‘’‚‛]/g, "'")
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }

  private static referenceEffectiveProperty(p: ParagraphSnapshot, property: string): unknown {
    switch (property) {
      case 'reference.style.fontName': return p.fontName;
      case 'reference.style.fontSize': return p.fontSize;
      case 'reference.style.alignment': return normalizeAlignment(p.alignment);
      case 'reference.style.leftIndent': return this.referenceStyleLeftIndent(p.referenceStyle);
      case 'reference.style.hangingIndent':
        return p.firstLineIndent === null || p.firstLineIndent === undefined
          ? null
          : Math.max(0, -p.firstLineIndent);
      default: return undefined;
    }
  }

  private static referenceRunPropertyValues(p: ParagraphSnapshot, property: string): unknown[] | null {
    if (property !== 'reference.style.fontName' && property !== 'reference.style.fontSize') return null;
    const runs = (p.effectiveRuns || []).filter(run => run.text.trim().length > 0);
    if (!runs.length) return null;
    return property === 'reference.style.fontName'
      ? runs.map(run => run.fontName)
      : runs.map(run => run.fontSize);
  }

  private static paragraphProperty(p: ParagraphSnapshot, property: string): unknown {
    const style = p.referenceStyle;
    const numbering = p.referenceNumbering;
    switch (property) {
      case 'font.name': return p.fontName;
      case 'font.size': return p.fontSize;
      case 'font.bold': return p.bold;
      case 'font.italic': return p.italic;
      case 'paragraph.alignment': return normalizeAlignment(p.alignment);
      case 'paragraph.firstLineIndent': return p.firstLineIndent;
      case 'paragraph.leftIndent': return p.leftIndent;
      case 'paragraph.rightIndent': return p.rightIndent;
      case 'paragraph.spaceBefore': return p.spaceBefore;
      case 'paragraph.spaceAfter': return p.spaceAfter;
      case 'paragraph.lineSpacing': return p.lineSpacing;
      case 'reference.style': {
        const name = style?.name || p.style;
        const next = style?.nextParagraphStyle || style?.name || p.style;
        return /^reference$/i.test(name || '')
          && /^paragraph$/i.test(style?.type || '')
          && /^paragraph$/i.test(style?.basedOn || '')
          && /^reference$/i.test(next || '');
      }
      case 'reference.style.fontName': return style?.fontName ?? null;
      case 'reference.style.fontSize': return style?.fontSize ?? null;
      case 'reference.style.alignment': return style?.alignment ? normalizeAlignment(style.alignment) : null;
      case 'reference.style.leftIndent': return this.referenceStyleLeftIndent(style);
      case 'reference.style.hangingIndent': return style?.hangingIndent ?? null;
      case 'reference.list.level': return numbering?.level ?? (p.isListItem ? 1 : null);
      case 'reference.list.numberPosition': return numbering?.numberPositionPt
        ?? (p.isListItem && p.leftIndent !== null && p.firstLineIndent !== null
          ? p.leftIndent + p.firstLineIndent
          : null);
      case 'reference.list.textPosition': return numbering?.textPositionPt ?? (p.isListItem ? p.leftIndent : null);
      default: return undefined;
    }
  }

  private static referenceStyleLeftIndent(style: ParagraphSnapshot['referenceStyle']): number | null {
    if (!style || style.leftIndent === null || style.leftIndent === undefined) return null;
    // Parsed DOCX styles expose the raw total left position.  Subtracting the
    // hanging distance yields the style's first-line/user-facing left anchor.
    // Synthetic/live snapshots that already expose the normalized value (for
    // example left=0, hanging=21.6) retain that value rather than turning it
    // into a negative position.
    if (style.leftIndent === 0) return 0;
    if (style.hangingIndent !== null && style.hangingIndent !== undefined
      && Math.abs(style.leftIndent - style.hangingIndent) <= 0.75) {
      return 0;
    }
    return style.leftIndent;
  }

  private static figureNumberSequence(targets: ParagraphSnapshot[]): {
    failures: ParagraphSnapshot[];
    missing: number[];
    duplicates: number[];
    mismatches: string[];
  } {
    const failures: ParagraphSnapshot[] = [];
    const missing = new Set<number>();
    const duplicates = new Set<number>();
    const seen = new Set<number>();
    const mismatches: string[] = [];
    for (const [position, paragraph] of targets.entries()) {
      const expected = position + 1;
      const actual = parseCaptionNumber(paragraph.text, 'figure');
      if (actual === null || !Number.isFinite(actual)) {
        failures.push(paragraph);
        mismatches.push(`Paragraph ${paragraph.index + 1}: expected ${expected}, detected unavailable`);
        continue;
      }
      if (seen.has(actual)) duplicates.add(actual);
      if (actual !== expected) {
        failures.push(paragraph);
        mismatches.push(`Paragraph ${paragraph.index + 1}: expected ${expected}, detected ${actual}`);
      }
      seen.add(actual);
    }
    if (seen.size) {
      const maximum = Math.max(...seen);
      for (let number = 1; number <= maximum; number += 1) {
        if (!seen.has(number)) missing.add(number);
      }
    }
    return {
      failures,
      missing: [...missing].sort((a, b) => a - b),
      duplicates: [...duplicates].sort((a, b) => a - b),
      mismatches
    };
  }

  private static missingSequenceNumbers(numbers: number[]): number[] {
    const present = new Set(numbers.filter(number => Number.isFinite(number)));
    if (!present.size) return [];
    const maximum = Math.max(...present);
    const missing: number[] = [];
    for (let number = 1; number <= maximum; number += 1) {
      if (!present.has(number)) missing.push(number);
    }
    return missing;
  }

  private static runPropertyValues(p: ParagraphSnapshot, property: string): unknown[] | null {
    if (!p.effectiveRuns?.length) return null;
    const runs = property === 'font.italic'
      ? this.runsRelevantToAuthorItalic(p)
      : p.classification === 'authorName'
        && (property === 'font.name' || property === 'font.size')
        ? p.effectiveRuns.filter(run =>
          run.text.trim().length > 0
          && !isRaisedAssociationRun(run)
        )
        : p.effectiveRuns;
    if (!runs.length) return null;
    switch (property) {
      case 'font.name': return runs.map(run => run.fontName);
      case 'font.size': return runs.map(run => run.fontSize);
      case 'font.bold': return runs.map(run => run.bold);
      case 'font.italic': return runs.map(run => run.italic);
      default: return null;
    }
  }

  private static runsRelevantToAuthorItalic(p: ParagraphSnapshot): NonNullable<ParagraphSnapshot['effectiveRuns']> {
    if (p.classification !== 'authorAffiliation' && p.classification !== 'authorEmail') {
      return p.effectiveRuns || [];
    }

    const normalizedText = this.normalizeLogicalLineBreaks(p.text);
    const bodySpans: Array<[number, number]> = [];
    if (p.classification === 'authorAffiliation') {
      const match = normalizedText.match(/^(\s*)(\d+(?:\s*[,;]\s*\d+)*)(\s*)(.*)$/);
      if (match && match[4].trim()) {
        const bodyStart = match[1].length + match[2].length + match[3].length;
        bodySpans.push([bodyStart, normalizedText.length]);
      }
    } else {
      for (const line of this.logicalLines(normalizedText)) {
        const match = line.text.match(/^(\s*)([A-Za-z]\))(\s*)(.*)$/);
        if (!match || !match[4].trim()) continue;
        const bodyStart = line.start + match[1].length + match[2].length + match[3].length;
        bodySpans.push([bodyStart, line.start + line.text.length]);
      }
    }

    // A missing association marker is handled by the separate superscript
    // rule. If there is no marker span, retain the historical fallback so a
    // plain affiliation/email paragraph can still be checked for italics.
    if (!bodySpans.length) return p.effectiveRuns || [];

    const relevantRuns: NonNullable<ParagraphSnapshot['effectiveRuns']> = [];
    let offset = 0;
    for (const run of p.effectiveRuns || []) {
      const runText = this.normalizeLogicalLineBreaks(run.text);
      const runStart = offset;
      const runEnd = runStart + runText.length;
      for (const [spanStart, spanEnd] of bodySpans) {
        const segmentStart = Math.max(runStart, spanStart);
        const segmentEnd = Math.min(runEnd, spanEnd);
        if (segmentEnd <= segmentStart) continue;
        relevantRuns.push({
          ...run,
          text: runText.slice(segmentStart - runStart, segmentEnd - runStart)
        });
      }
      offset = runEnd;
    }
    return relevantRuns;
  }

  private static matches(detected: unknown, expected: unknown, property: string): boolean {
    if (typeof expected === 'number') return typeof detected === 'number' && approxEqual(detected, expected, property.startsWith('layout.') ? 0.02 : 0.75);
    if (typeof expected === 'boolean') return detected === expected;
    if (typeof expected === 'string') {
      if (property === 'font.name') return typeof detected === 'string' && detected.toLowerCase() === expected.toLowerCase();
      if (property === 'paragraph.alignment') return normalizeAlignment(String(detected)) === normalizeAlignment(expected);
      return String(detected).trim().toLowerCase() === expected.trim().toLowerCase();
    }
    return detected === expected;
  }

  private static expected(rule: RuleDefinition): string {
    if (rule.operator === 'required') return 'Present';
    if (rule.operator === 'forbidden') return 'Absent';
    return displayValue(rule.expected, rule.unit);
  }

  private static finding(
    rule: RuleDefinition,
    status: Finding['status'],
    paragraphIndices: number[],
    expected: string,
    detected: string,
    message: string,
    fixParagraphIndices?: number[],
    reviewKind?: ReviewKind,
    fixTableIndices?: number[],
    autoFixableOverride?: boolean,
    fixTableHeaderRowCounts?: number[],
    fixTableCaptionRowIndices?: Array<number | null>,
    fixTableHeaderStartRowIndices?: Array<number | null>,
    fixTableHeaderEndRowIndices?: Array<number | null>
  ): Finding {
    return {
      ruleId: rule.id,
      ruleName: rule.name,
      category: rule.category,
      target: rule.target,
      status,
      severity: rule.severity,
      expected,
      detected,
      message,
      paragraphIndices,
      fixParagraphIndices,
      autoFixable: autoFixableOverride !== undefined
        ? autoFixableOverride && status === 'failed'
        : rule.autoFixable && status === 'failed' && (
          paragraphIndices.length > 0 || rule.target === 'document'
        ),
      fixTableIndices,
      fixTableHeaderRowCounts,
      fixTableCaptionRowIndices,
      fixTableHeaderStartRowIndices,
      fixTableHeaderEndRowIndices,
      property: rule.property,
      rawExpected: rule.expected,
      evidence: rule.evidence,
      note: rule.note,
      reviewKind,
      assessmentRole: rule.assessmentRole || 'compliance'
    };
  }
}
