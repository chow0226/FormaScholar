import {
  CitationOccurrence,
  CrossReferenceSnapshot,
  DefinedObject,
  DocumentBlockSnapshot,
  FieldTextSpan,
  ObjectPlacementRelation,
  ParagraphSnapshot
} from '../core/models';
import { parseCaptionNumber } from '../core/CaptionNumberParser';

function expandNumericList(value: string): number[] {
  const numbers: number[] = [];
  for (const token of value.split(',')) {
    const match = token.trim().match(/^(\d+)\s*(?:[-\u2013]\s*(\d+))?$/);
    if (!match) continue;
    const start = Number(match[1]);
    const end = match[2] ? Number(match[2]) : start;
    if (!Number.isFinite(start) || !Number.isFinite(end) || end < start || end - start > 500) continue;
    for (let n = start; n <= end; n += 1) numbers.push(n);
  }
  return [...new Set(numbers)];
}

function expandObjectNumericList(value: string): number[] {
  const normalized = value
    .replace(/,\s*and\s+/gi, ',')
    .replace(/\band\b/gi, ',');
  const numbers: number[] = [];

  for (const token of normalized.split(',')) {
    const match = token.trim().match(/^(\d+)(?:[A-Za-z]|\s*\([A-Za-z]\))?(?:\s*[-\u2013]\s*(\d+)(?:[A-Za-z]|\s*\([A-Za-z]\))?)?$/);
    if (!match) continue;
    const start = Number(match[1]);
    const end = match[2] ? Number(match[2]) : start;
    if (!Number.isFinite(start) || !Number.isFinite(end) || end < start || end - start > 500) continue;
    for (let n = start; n <= end; n += 1) numbers.push(n);
  }

  return [...new Set(numbers)];
}

export function extractReferenceCitations(text: string): number[][] {
  const citations: number[][] = [];
  for (const match of text.matchAll(/\[\s*([0-9]+(?:\s*[-\u2013]\s*[0-9]+)?(?:\s*,\s*[0-9]+(?:\s*[-\u2013]\s*[0-9]+)?)*)\s*\]/g)) {
    const numbers = expandNumericList(match[1]);
    if (numbers.length) citations.push(numbers);
  }
  return citations;
}

function extractObjectCitations(text: string, kind: 'figure' | 'table'): number[][] {
  const prefix = kind === 'figure' ? '(?:figure|figures|fig\\.|figs\\.)' : '(?:table|tables)';
  const result: number[][] = [];
  const separator = '(?:\\s*,\\s*(?:and\\s+)?|\\s+and\\s+|\\s*[-\\u2013]\\s*)';
  const objectNumber = '[0-9]+(?:[A-Z]|\\s*\\([A-Z]\\))?';
  const pattern = new RegExp(`\\b${prefix}\\s+(${objectNumber}(?:${separator}${objectNumber})*)`, 'gi');
  for (const match of text.matchAll(pattern)) {
    const numbers = expandObjectNumericList(match[1]);
    if (numbers.length) result.push(numbers);
  }
  return result;
}

export function extractFigureCitationOccurrences(text: string, paragraphIndex: number, fieldTextSpans: FieldTextSpan[] = []): CitationOccurrence[] {
  const separator = '(?:\\s*,\\s*(?:and\\s+)?|\\s+and\\s+|\\s*[-\\u2013]\\s*)';
  const objectNumber = '[0-9]+(?:[A-Z]|\\s*\\([A-Z]\\))?';
  const pattern = new RegExp(
    `\\b(figure|figures|fig\\.|figs\\.)\\s+(${objectNumber}(?:${separator}${objectNumber})*)`,
    'gi'
  );
  const occurrenceByCitation = new Map<string, number>();
  const occurrences: CitationOccurrence[] = [];

  for (const match of text.matchAll(pattern)) {
    const prefix = match[1];
    const numbers = expandObjectNumericList(match[2]);
    if (!numbers.length || match.index === undefined) continue;

    const before = text.slice(0, match.index)
      .replace(/[\s"'“”‘’([{]+$/u, '');
    const abbreviationBeforeCitation = /(?:^|[\s"'“‘([{])(?:e\.g|i\.e|etc|cf|eq|ref|refs|dr|mr|mrs|ms|viz|vs)\.$/iu.test(before);
    const isSentenceStart = before.length === 0
      || /[!?]$/u.test(before)
      || (/\.$/u.test(before) && !abbreviationBeforeCitation);
    const plural = prefix.toLowerCase() === 'figures' || prefix.toLowerCase() === 'figs.';
    // The approved FormaScholar/AIP refinement always spells out Figure in
    // manuscript text. Sentence position is retained as diagnostic metadata,
    // but no longer changes the expected wording.
    const expectedPrefix = plural ? 'Figures' : 'Figure';
    const originalCitationText = match[0];
    const correctedCitationText = `${expectedPrefix}${originalCitationText.slice(prefix.length)}`;
    const citationEnd = match.index + originalCitationText.length;
    const isFieldBacked = fieldTextSpans.some(span => span.start < citationEnd && span.end > match.index!);
    const occurrenceIndex = occurrenceByCitation.get(originalCitationText) || 0;
    occurrenceByCitation.set(originalCitationText, occurrenceIndex + 1);

    occurrences.push({
      kind: 'figure',
      numbers,
      paragraphIndex,
      rawText: text,
      start: match.index,
      end: match.index + prefix.length,
      originalText: prefix,
      originalCitationText,
      correctedCitationText,
      isFieldBacked,
      isSentenceStart,
      expectedPrefix,
      occurrenceIndex
    });
  }

  return occurrences;
}

function definitionNumber(text: string, kind: 'reference' | 'figure' | 'table'): number | null {
  if (kind !== 'reference') return parseCaptionNumber(text, kind);
  const match = text.trim().match(/^(?:\[\s*)?(\d+)(?:\s*\])?[.)]/);
  return match ? Number(match[1]) : null;
}

function referenceDefinitionNumber(paragraph: ParagraphSnapshot): number | null {
  const textNumber = definitionNumber(paragraph.text, 'reference');
  if (textNumber !== null) return textNumber;

  const listNumber = (paragraph.listString || '').trim().match(/^(\d+)[.)]?$/);
  return listNumber ? Number(listNumber[1]) : null;
}

function isPresent<T>(value: T | null): value is T {
  return value !== null;
}

function citationOccurrences(
  paragraphs: ParagraphSnapshot[],
  referencesHeadingIndex: number | null,
  kind: 'reference' | 'figure' | 'table'
): CitationOccurrence[] {
  return paragraphs
    .filter(p => referencesHeadingIndex === null || p.index < referencesHeadingIndex)
    .filter(p => {
      if (kind === 'reference') return true;
      return p.tableNestingLevel === 0
        && p.classification !== 'figureCaption'
        && p.classification !== 'tableCaption';
    })
    .flatMap(p => {
      if (kind === 'figure') return extractFigureCitationOccurrences(p.text, p.index, p.fieldTextSpans);
      const groups = kind === 'reference'
        ? extractReferenceCitations(p.text)
        : extractObjectCitations(p.text, kind);
      return groups.map(numbers => ({ kind, numbers, paragraphIndex: p.index, rawText: p.text }));
    });
}

function placementFor(
  definitions: DefinedObject[],
  blocks: DocumentBlockSnapshot[],
  kind: 'figure' | 'table'
): ObjectPlacementRelation[] {
  if (!blocks.length) {
    return definitions.map(definition => ({
      kind,
      number: definition.number,
      captionParagraphIndex: definition.paragraphIndex,
      objectBlockIndex: null,
      passed: null,
      confidence: 'low'
    }));
  }

  const captionBlocks = definitions
    .map(definition => {
      const block = blocks.find(candidate => candidate.paragraphIndex === definition.paragraphIndex);
      return block ? { definition, block } : null;
    })
    .filter((entry): entry is { definition: DefinedObject; block: DocumentBlockSnapshot } => entry !== null)
    .sort((a, b) => a.block.blockIndex - b.block.blockIndex);

  const figureObjectConfidence = (block: DocumentBlockSnapshot | undefined): 'high' | 'medium' | 'low' => {
    if (!block || (block.drawingCount || 0) === 0) return 'low';
    if (block.kind === 'paragraph' && block.hasInlineDrawing && block.hasAnchoredDrawing) return 'low';
    if (block.kind === 'paragraph' && block.hasAnchoredDrawing) return 'low';
    if (block.kind === 'paragraph' && block.hasInlineDrawing) return 'high';
    if (block.kind === 'table') return 'medium';
    return 'low';
  };

  const isTableObject = (block: DocumentBlockSnapshot | undefined) =>
    block?.kind === 'table' && block.isEquationLayout !== true;
  const isFigureObject = (block: DocumentBlockSnapshot | undefined) => (block?.drawingCount || 0) > 0;
  const confidenceForFigure = (block: DocumentBlockSnapshot) => figureObjectConfidence(block);
  const relation = (
    definition: DefinedObject,
    object: DocumentBlockSnapshot | undefined,
    passed: boolean,
    confidence: 'high' | 'medium' | 'low'
  ): ObjectPlacementRelation => ({
    kind,
    number: definition.number,
    captionParagraphIndex: definition.paragraphIndex,
    objectBlockIndex: object?.blockIndex ?? null,
    passed: confidence === 'low' ? null : passed,
    confidence
  });
  const unknown = (definition: DefinedObject): ObjectPlacementRelation => ({
    kind,
    number: definition.number,
    captionParagraphIndex: definition.paragraphIndex,
    objectBlockIndex: null,
    passed: null,
    confidence: 'low'
  });

  return definitions.map(definition => {
    if (kind === 'table' && definition.containingTableIndex !== null && definition.containingTableIndex !== undefined) {
      const containingTable = blocks.find(block =>
        block.kind === 'table' && block.tableIndex === definition.containingTableIndex
      );
      if (containingTable) {
        // Embedded captions are valid only in the first logical row.  Keep a
        // high-confidence association for a late caption so the placement
        // rule can report a real failure, but do not let it redefine the
        // table's structural caption/header rows.
        const validEmbeddedCaption = definition.containingTableRowIndex === 0;
        return relation(definition, containingTable, validEmbeddedCaption, 'high');
      }
    }
    const currentEntry = captionBlocks.find(entry =>
      entry.definition.paragraphIndex === definition.paragraphIndex
    );
    if (!currentEntry) return unknown(definition);

    const currentPosition = currentEntry.block.blockIndex;
    const captionIndex = captionBlocks.indexOf(currentEntry);
    const previousCaptionPosition = captionIndex > 0
      ? captionBlocks[captionIndex - 1].block.blockIndex
      : -1;
    const nextCaptionPosition = captionIndex + 1 < captionBlocks.length
      ? captionBlocks[captionIndex + 1].block.blockIndex
      : Number.POSITIVE_INFINITY;
    const between = (start: number, end: number, predicate: (block: DocumentBlockSnapshot) => boolean) =>
      blocks.filter(block => block.blockIndex > start && block.blockIndex < end && predicate(block));

    if (kind === 'figure') {
      const before = between(previousCaptionPosition, currentPosition, isFigureObject);
      const after = between(currentPosition, nextCaptionPosition, isFigureObject);

      // The previous caption creates a boundary for the preceding figure.
      // This pairs consecutive object/caption blocks without comparing the
      // current figure to the next figure's object.
      if (before.length === 1) {
        if (captionIndex === 0 && nextCaptionPosition === Number.POSITIVE_INFINITY && after.length > 0) return unknown(definition);
        return relation(definition, before[0], true, confidenceForFigure(before[0]));
      }
      if (before.length > 1) return unknown(definition);

      // No object before the caption: a single object in the following
      // caption window is a confidently misplaced caption.
      if (after.length === 1) return relation(definition, after[0], false, confidenceForFigure(after[0]));
      return unknown(definition);
    }

    const before = between(previousCaptionPosition, currentPosition, isTableObject);
    const after = between(currentPosition, nextCaptionPosition, isTableObject);
    const firstAfter = after[0];
    const paragraphsBeforeFirstTable = firstAfter
      ? blocks.filter(block => block.blockIndex > currentPosition && block.blockIndex < firstAfter.blockIndex)
      : [];
    const immediatelyFollowedByTable = Boolean(firstAfter)
      && paragraphsBeforeFirstTable.every(block =>
        block.kind === 'paragraph' && (block.text || '').trim().length === 0
      );
    if (immediatelyFollowedByTable && firstAfter) {
      if (captionIndex === 0 && nextCaptionPosition === Number.POSITIVE_INFINITY && before.length > 0) return unknown(definition);
      return relation(definition, firstAfter, true, 'high');
    }
    if (after.length === 1) {
      if (captionIndex === 0 && nextCaptionPosition === Number.POSITIVE_INFINITY && before.length > 0) return unknown(definition);
      return relation(definition, after[0], true, 'high');
    }
    if (after.length > 1) return unknown(definition);

    if (before.length === 1) return relation(definition, before[0], false, 'high');
    return unknown(definition);
  });
}

export function analyzeCrossReferences(
  paragraphs: ParagraphSnapshot[],
  blocks: DocumentBlockSnapshot[] = [],
  referencesHeadingIndex: number | null = null
): CrossReferenceSnapshot {
  const referenceDefinitions = paragraphs
    .filter(p => p.classification === 'reference')
    .map(p => {
      const number = referenceDefinitionNumber(p);
      return number === null ? null : { kind: 'reference' as const, number, paragraphIndex: p.index };
    })
    .filter(isPresent);
  const figureDefinitions = paragraphs
    .filter(p => p.classification === 'figureCaption')
    .map(p => {
      const number = definitionNumber(p.text, 'figure');
      return number === null ? null : { kind: 'figure' as const, number, paragraphIndex: p.index };
    })
    .filter(isPresent);
  const tableDefinitions = paragraphs
    .filter(p => p.classification === 'tableCaption')
    .map(p => {
      const number = definitionNumber(p.text, 'table');
      return number === null ? null : {
        kind: 'table' as const,
        number,
        paragraphIndex: p.index,
        containingTableIndex: p.containingTableIndex ?? null,
        containingTableRowIndex: p.containingTableRowIndex ?? null
      };
    })
    .filter(isPresent);

  const referenceCitations = citationOccurrences(paragraphs, referencesHeadingIndex, 'reference');
  const figureCitations = citationOccurrences(paragraphs, referencesHeadingIndex, 'figure');
  const tableCitations = citationOccurrences(paragraphs, referencesHeadingIndex, 'table');
  const referenceNumbers = new Set(referenceDefinitions.map(item => item.number));
  const citedReferences = new Set(referenceCitations.flatMap(item => item.numbers));
  const figureNumbers = new Set(figureDefinitions.map(item => item.number));
  const citedFigures = new Set(figureCitations.flatMap(item => item.numbers));
  const tableNumbers = new Set(tableDefinitions.map(item => item.number));
  const citedTables = new Set(tableCitations.flatMap(item => item.numbers));

  const sequenceViolations: CitationOccurrence[] = [];
  let nextExpected = 1;
  const seen = new Set<number>();
  for (const citation of referenceCitations) {
    let violation = false;
    for (const number of citation.numbers) {
      // Invalid identifiers belong exclusively to the validity finding. Do
      // not make them create a second, misleading sequence failure.
      if (!referenceNumbers.has(number)) continue;
      if (seen.has(number)) continue;
      if (number !== nextExpected) violation = true;
      seen.add(number);
      while (seen.has(nextExpected)) nextExpected += 1;
    }
    if (violation) sequenceViolations.push(citation);
  }

  const figureCitationSequenceViolations: CitationOccurrence[] = [];
  let nextFigureExpected = 1;
  const seenFigures = new Set<number>();
  for (const citation of figureCitations) {
    let violation = false;
    for (const number of citation.numbers) {
      if (!figureNumbers.has(number) || seenFigures.has(number)) continue;
      if (number !== nextFigureExpected) violation = true;
      seenFigures.add(number);
      while (seenFigures.has(nextFigureExpected)) nextFigureExpected += 1;
    }
    if (violation) figureCitationSequenceViolations.push(citation);
  }

  const paragraphBlockMap = new Map<number, number>();
  for (const block of blocks) {
    if (block.kind === 'paragraph' && block.paragraphIndex !== undefined) {
      paragraphBlockMap.set(block.paragraphIndex, block.blockIndex);
    }
  }
  const figurePlacement = placementFor(figureDefinitions, blocks, 'figure');
  const figurePriorCitationFailures: DefinedObject[] = [];
  const figurePriorCitationUnknown: DefinedObject[] = [];
  for (const definition of figureDefinitions) {
    const placement = figurePlacement.find(item =>
      item.number === definition.number &&
      item.captionParagraphIndex === definition.paragraphIndex
    );
    const citations = figureCitations.filter(citation => citation.numbers.includes(definition.number));
    if (!placement || placement.confidence !== 'high' || placement.objectBlockIndex === null) {
      figurePriorCitationUnknown.push(definition);
      continue;
    }
    if (!citations.length) {
      figurePriorCitationFailures.push(definition);
      continue;
    }
    const mappedCitations = citations
      .map(citation => paragraphBlockMap.get(citation.paragraphIndex))
      .filter((blockIndex): blockIndex is number => blockIndex !== undefined);
    if (!mappedCitations.length) {
      figurePriorCitationUnknown.push(definition);
    } else if (!mappedCitations.some(blockIndex => blockIndex < placement.objectBlockIndex!)) {
      figurePriorCitationFailures.push(definition);
    }
  }

  return {
    referenceDefinitions,
    referenceCitations,
    figureDefinitions,
    figureCitations,
    tableDefinitions,
    tableCitations,
    referenceCoverageMissing: [...referenceNumbers].filter(number => !citedReferences.has(number)),
    invalidReferenceCitations: referenceCitations.filter(citation => citation.numbers.some(number => !referenceNumbers.has(number))),
    referenceSequenceViolations: sequenceViolations,
    figureCoverageMissing: [...figureNumbers].filter(number => !citedFigures.has(number)),
    tableCoverageMissing: [...tableNumbers].filter(number => !citedTables.has(number)),
    referenceCitationStyle: referenceCitations.length ? 'square_bracket_numeric' : 'unknown',
    figurePlacement,
    tablePlacement: placementFor(tableDefinitions, blocks, 'table'),
    figurePriorCitationFailures,
    figurePriorCitationUnknown,
    figureCitationSequenceViolations
  };
}
