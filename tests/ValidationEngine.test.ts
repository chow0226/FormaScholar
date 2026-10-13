import { describe, expect, it, vi } from 'vitest';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import strategyJson from '../src/strategies/aip-2023.json';
import { analyzeAffiliationEmailBlankLine, analyzeFigureSpaceAbove, analyzeFrontMatterTransitions, analyzeTableCaptionTableGaps, findExtraBlankBodyParagraphIndices, findFigureCaptionBodyGapIssues, ValidationEngine } from '../src/core/ValidationEngine';
import { CorrectionService, isVerifiedSingleLineSpacing, verifySingleLineSpacing } from '../src/core/CorrectionService';
import { DocumentSnapshot, Finding, ParagraphSnapshot, SectionLayoutSnapshot, StrategyDefinition } from '../src/core/models';
import { aggregateBooleanState, collapseRepeatedOrdinarySpaces, extractCaptionPrefix, extractFigureCaptionLabel, groupRaisedAssociationMarkerRuns, hasExplicitCaptionPrefix, tableBoundariesToRestore, toSafeTitleCase, WordDocumentAdapter } from '../src/word/WordDocumentAdapter';
import { DocxInspector, hasPageNumberField, inspectBlankParagraph, leadingManualLineBreakCount, parseDocumentBlocks, parseParagraphOoxml, repairParagraphIndentOoxml } from '../src/ooxml/DocxInspector';
import { classifyFieldInstruction, findPageNumberFields, findWordFields, isTableSequenceFieldInstruction } from '../src/ooxml/WordFieldAnalyzer';
import { resolveEffectiveFormatting } from '../src/ooxml/EffectiveFormattingResolver';
import { analyzeCrossReferences, extractFigureCitationOccurrences, extractReferenceCitations } from '../src/ooxml/CrossReferenceAnalyzer';
import { documentFingerprint } from '../src/core/DocumentFingerprint';
import { filterFindings } from '../src/taskpane/FindingsFilter';
import { ApplicabilityAnalyzer } from '../src/core/ApplicabilityAnalyzer';
import { parseCaptionNumber, romanNumeralToNumber } from '../src/core/CaptionNumberParser';
import { canonicalSectionsForHeading, normalizeSectionHeading } from '../src/core/SectionOntology';

const strategy = strategyJson as StrategyDefinition;

function paragraph(overrides: Partial<ParagraphSnapshot> = {}): ParagraphSnapshot {
  return {
    index: 0,
    text: 'Example paragraph',
    style: 'Normal',
    styleBuiltIn: 'Normal',
    alignment: 'justified',
    firstLineIndent: 14.4,
    leftIndent: 0,
    rightIndent: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: 12,
    lineSpacingMultiple: 1,
    lineSpacingRule: 'auto',
    tableNestingLevel: 0,
    fontName: 'Times New Roman',
    fontSize: 10,
    bold: false,
    italic: false,
    allCaps: false,
    isListItem: false,
    listString: null,
    abstractPrefixBold: null,
    abstractBodyBold: false,
    abstractBodyMixed: false,
    figurePrefixBold: null,
    figurePrefixAllCaps: null,
    tablePrefixBold: null,
    tablePrefixAllCaps: null,
    classification: 'body',
    ...overrides
  };
}

function snapshot(paragraphs: ParagraphSnapshot[], sections: SectionLayoutSnapshot[] = [{
  index: 0,
  pageWidthInches: 8.5,
  pageHeightInches: 11,
  marginTopInches: 1,
  marginRightInches: 1,
  marginBottomInches: 1,
  marginLeftInches: 1,
  columns: 1
}]): DocumentSnapshot {
  return {
    paragraphs,
    layout: {
      ...sections[sections.length - 1],
      sections,
      hasHeaderReference: false,
      hasFooterReference: false,
      containsPageNumberField: false,
      paragraphMappingAligned: true,
      ooxmlAvailable: true,
      ooxmlParser: 'fallback'
    },
    referencesHeadingIndex: null,
    abstractIndex: paragraphs.findIndex(p => p.classification === 'abstract')
  };
}

function validateRule(ruleId: string, doc: DocumentSnapshot) {
  const selected = strategy.rules.find(rule => rule.id === ruleId);
  if (!selected) throw new Error(`Missing strategy rule ${ruleId}`);
  return ValidationEngine.validate({ ...strategy, rules: [selected] }, doc).findings[0];
}

function correctionFinding(overrides: Partial<Finding>): Finding {
  return {
    ruleId: 'BODY_FONT_001',
    ruleName: 'Body font',
    category: 'Body',
    target: 'body',
    status: 'failed',
    severity: 'minor',
    expected: 'Times New Roman',
    detected: 'Calibri',
    message: 'Formatting does not match.',
    paragraphIndices: [0],
    autoFixable: true,
    property: 'font.name',
    rawExpected: 'Times New Roman',
    evidence: 'Test evidence.',
    ...overrides
  };
}

describe('AIP deterministic validation', () => {
  it('preserves mixed author-body boolean state for run-level fallback', () => {
    expect(aggregateBooleanState([])).toBeNull();
    expect(aggregateBooleanState([true])).toBe(true);
    expect(aggregateBooleanState([true, true])).toBe(true);
    expect(aggregateBooleanState([false, false])).toBe(false);
    expect(aggregateBooleanState([false, true])).toBeNull();
  });

  it('uses 1.00 in bottom margin in every section', () => {
    const valid = validateRule('MARGIN_BOTTOM_001', snapshot([], [
      { index: 0, marginBottomInches: 1 },
      { index: 1, marginBottomInches: 1 }
    ]));
    const invalid = validateRule('MARGIN_BOTTOM_001', snapshot([], [
      { index: 0, marginBottomInches: 1 },
      { index: 1, marginBottomInches: 0.75 }
    ]));

    expect(valid.status).toBe('passed');
    expect(invalid.status).toBe('failed');
    expect(invalid.detected).toContain('Section 2 = 0.75 in');
    expect(invalid.autoFixable).toBe(true);
  });

  it('offers safe fixes for bottom margins and forbidden page-number fields', () => {
    const pageNumberDocument = snapshot([]);
    pageNumberDocument.layout.containsPageNumberField = true;
    expect(validateRule('MARGIN_BOTTOM_001', snapshot([], [{ index: 0, marginBottomInches: 0.8 }])).autoFixable).toBe(true);
    expect(validateRule('PAGE_NUMBER_FORBIDDEN_001', pageNumberDocument).autoFixable).toBe(true);
    expect(validateRule('PAGE_NUMBER_FORBIDDEN_001', pageNumberDocument).status).toBe('failed');
  });

  it('recognizes PAGE, NUMPAGES, and SECTIONPAGES fields without matching ordinary prose', () => {
    expect(hasPageNumberField('<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>')).toBe(true);
    expect(hasPageNumberField('<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>10</w:t></w:r></w:fldSimple>')).toBe(true);
    expect(hasPageNumberField('<w:fldSimple w:instr=" SECTIONPAGES "><w:r><w:t>2</w:t></w:r></w:fldSimple>')).toBe(true);
    expect(hasPageNumberField('<w:p><w:r><w:t>Page 3 of 10</w:t></w:r></w:p>')).toBe(false);
  });

  it('does not mistake Zotero page metadata for a PAGE field', () => {
    const zotero = `
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve">ADDIN ZOTERO_ITEM CSL_CITATION {&quot;itemData&quot;:{&quot;page&quot;:&quot;8-12&quot;}}</w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    expect(classifyFieldInstruction('ADDIN ZOTERO_ITEM CSL_CITATION {"page":"8-12"}')).toBe('OTHER');
    expect(hasPageNumberField(zotero)).toBe(false);
    expect(findPageNumberFields(zotero)).toEqual([]);
  });

  it('reconstructs fragmented complex PAGE instructions', () => {
    const xml = `
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText>PA</w:instrText></w:r>
      <w:r><w:instrText>GE</w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>4</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const fields = findWordFields(xml);
    expect(fields).toHaveLength(1);
    expect(fields[0].instruction).toBe('PAGE');
    expect(fields[0].kind).toBe('PAGE');
    expect(hasPageNumberField(xml)).toBe(true);
  });

  it('has dedicated correction paths for section margins and page-number fields', async () => {
    const source = await readFile('src/word/WordDocumentAdapter.ts', 'utf8');
    expect(source).toContain('static async applySectionMargin(');
    expect(source).toContain('static async removePageNumberFields()');
    expect(source).toContain("field.delete()");
    expect(source).toContain('section.pageSetup.bottomMargin = points');
  });

  it('does not score caption presence when no object candidates exist', () => {
    expect(validateRule('FIGURE_CAPTION_PRESENCE_001', snapshot([])).status).toBe('not_applicable');
    expect(validateRule('TABLE_CAPTION_PRESENCE_001', snapshot([])).status).toBe('not_applicable');
  });

  it('does not treat equation-layout tables as manuscript table objects', () => {
    const doc = snapshot([]);
    doc.layout.documentBlocks = [{
      blockIndex: 0,
      kind: 'table',
      tableIndex: 0,
      isEquationLayout: true,
      drawingCount: 0,
      text: 'x = y (1)'
    }];

    const features = ApplicabilityAnalyzer.analyze(doc);

    expect(features.hasTableObjects).toBe(false);
    expect(features.hasTables).toBe(false);
    expect(features.tableObjectCount).toBe(0);
    expect(validateRule('TABLE_CAPTION_PRESENCE_001', doc).status).toBe('not_applicable');
  });

  it('fingerprints paragraph text and table structure for correction safety', () => {
    const original = snapshot([paragraph({ index: 0, text: 'Original text' })]);
    const changedText = snapshot([paragraph({ index: 0, text: 'Changed text' })]);
    const changedTable = snapshot([paragraph({ index: 0, text: 'Original text' })]);
    const changedMembership = snapshot([paragraph({ index: 0, text: 'Original text', containingTableIndex: 0, containingTableRowIndex: 1, tableNestingLevel: 1 })]);
    const changedClassification = snapshot([paragraph({ index: 0, text: 'Original text', classification: 'heading1' })]);
    const changedFieldSpan = snapshot([paragraph({ index: 0, text: 'Original text', hasField: true, fieldTextSpans: [{ start: 2, end: 3 }] })]);
    const movedFieldSpan = snapshot([paragraph({ index: 0, text: 'Original text', hasField: true, fieldTextSpans: [{ start: 5, end: 6 }] })]);
    changedTable.layout.documentBlocks = [{ blockIndex: 0, kind: 'table', tableIndex: 0, drawingCount: 0 }];
    expect(documentFingerprint(original)).not.toBe(documentFingerprint(changedText));
    expect(documentFingerprint(original)).not.toBe(documentFingerprint(changedTable));
    expect(documentFingerprint(original)).not.toBe(documentFingerprint(changedMembership));
    expect(documentFingerprint(original)).not.toBe(documentFingerprint(changedClassification));
    expect(documentFingerprint(original)).not.toBe(documentFingerprint(changedFieldSpan));
    expect(documentFingerprint(changedFieldSpan)).not.toBe(documentFingerprint(movedFieldSpan));
    expect(documentFingerprint(original)).toBe(documentFingerprint(snapshot([paragraph({ index: 0, text: '  Original   text ' })])));
  });

  it('fingerprints field kind changes even when visible text and span offsets stay the same', () => {
    const sequenceField = snapshot([paragraph({
      index: 0,
      text: 'TABLE 1. Results',
      hasField: true,
      fieldTextSpans: [{ start: 6, end: 7, kind: 'tableSequence' }]
    })]);
    const referenceField = snapshot([paragraph({
      index: 0,
      text: 'TABLE 1. Results',
      hasField: true,
      fieldTextSpans: [{ start: 6, end: 7, kind: 'other' }]
    })]);

    expect(documentFingerprint(sequenceField)).not.toBe(documentFingerprint(referenceField));
  });

  it('restores table boundaries only when removing insideH requires restoration', () => {
    expect(tableBoundariesToRestore(false, true, true)).toEqual({ captionHeader: false, headerBody: false });
    expect(tableBoundariesToRestore(true, true, true)).toEqual({ captionHeader: true, headerBody: true });
    expect(tableBoundariesToRestore(true, false, true)).toEqual({ captionHeader: false, headerBody: true });
  });

  it('requires and can correct author-email italic formatting', () => {
    const pass = validateRule('AUTHOR_EMAIL_ITALIC_001', snapshot([
      paragraph({ classification: 'authorEmail', italic: true })
    ]));
    const fail = validateRule('AUTHOR_EMAIL_ITALIC_001', snapshot([
      paragraph({ classification: 'authorEmail', italic: false })
    ]));

    expect(pass.status).toBe('passed');
    expect(fail.status).toBe('failed');
    expect(fail.autoFixable).toBe(true);
  });

  it('describes the opposite italic direction when an imported rule expects non-italic text', () => {
    const baseRule = strategy.rules.find(rule => rule.id === 'AUTHOR_EMAIL_ITALIC_001');
    const affiliationBaseRule = strategy.rules.find(rule => rule.id === 'AUTHOR_AFFILIATION_ITALIC_001');
    if (!baseRule) throw new Error('Missing author email italic rule');
    if (!affiliationBaseRule) throw new Error('Missing author affiliation italic rule');
    const nonItalicRule = { ...baseRule, expected: false };
    const nonItalicAffiliationRule = { ...affiliationBaseRule, expected: false };
    const finding = ValidationEngine.validate(
      { ...strategy, rules: [nonItalicRule] },
      snapshot([paragraph({ classification: 'authorEmail', authorAssociationBodyItalic: true })])
    ).findings[0];
    const affiliationFinding = ValidationEngine.validate(
      { ...strategy, rules: [nonItalicAffiliationRule] },
      snapshot([paragraph({ classification: 'authorAffiliation', authorAssociationBodyItalic: true })])
    ).findings[0];

    expect(finding.status).toBe('failed');
    expect(finding.detected).toContain('Email body contains italic text');
    expect(finding.message).toContain('contain italic email text');
    expect(finding.detected).not.toContain('non-italic text');
    expect(affiliationFinding.status).toBe('failed');
    expect(affiliationFinding.detected).toContain('Affiliation body contains italic text');
    expect(affiliationFinding.message).toContain('contain italic affiliation text');
  });

  it('fails non-italic author validation when body runs are mixed', () => {
    const baseRule = strategy.rules.find(rule => rule.id === 'AUTHOR_EMAIL_ITALIC_001');
    if (!baseRule) throw new Error('Missing author email italic rule');
    const nonItalicRule = { ...baseRule, expected: false };
    const finding = ValidationEngine.validate(
      { ...strategy, rules: [nonItalicRule] },
      snapshot([paragraph({
        classification: 'authorEmail',
        authorAssociationBodyItalic: null,
        effectiveRuns: [
          { text: 'first@example.com', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, verticalAlign: null, sources: {} },
          { text: 'second@example.com', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: true, allCaps: false, verticalAlign: null, sources: {} }
        ]
      })])
    ).findings[0];

    expect(finding.status).toBe('failed');
    expect(finding.detected).toContain('italic text');
  });

  it('ignores leading superscript author markers when checking italic text', () => {
    const affiliationRuns = [
      { text: '1 ', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, verticalAlign: 'superscript' as const, sources: {} },
      { text: 'Faculty of Science', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: true, allCaps: false, verticalAlign: null, sources: {} }
    ];
    const emailRuns = [
      { text: 'a) ', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, verticalAlign: 'superscript' as const, sources: {} },
      { text: 'author@example.com', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: true, allCaps: false, verticalAlign: null, sources: {} }
    ];

    expect(validateRule('AUTHOR_AFFILIATION_ITALIC_001', snapshot([
      paragraph({ classification: 'authorAffiliation', text: '1 Faculty of Science', italic: null, effectiveRuns: affiliationRuns })
    ])).status).toBe('passed');
    expect(validateRule('AUTHOR_EMAIL_ITALIC_001', snapshot([
      paragraph({ classification: 'authorEmail', text: 'a) author@example.com', italic: null, effectiveRuns: emailRuns })
    ])).status).toBe('passed');
    expect(validateRule('AUTHOR_AFFILIATION_ITALIC_001', snapshot([
      paragraph({ classification: 'authorAffiliation', text: '1 Faculty of Science', italic: null, effectiveRuns: affiliationRuns.map(run => ({ ...run, italic: false })) })
    ])).status).toBe('failed');
    expect(validateRule('AUTHOR_AFFILIATION_ITALIC_001', snapshot([
      paragraph({
        classification: 'authorAffiliation',
        text: '1 Faculty of Science',
        italic: null,
        effectiveRuns: affiliationRuns.map(run => ({ ...run, verticalAlign: null }))
      })
    ])).status).toBe('passed');
  });

  it('groups fragmented raised author markers before correction searches', () => {
    const marker = (text: string) => ({
      text,
      fontName: 'Times New Roman',
      fontSize: 11,
      bold: false,
      italic: false,
      allCaps: false,
      verticalAlign: null,
      baselinePositionPt: 5,
      sources: {}
    });
    const groups = groupRaisedAssociationMarkerRuns([
      { ...marker('Panca'), baselinePositionPt: null, verticalAlign: null },
      marker('1'),
      marker(','),
      marker(' '),
      marker('c'),
      marker(')'),
    ]);

    expect(groups).toEqual([{ text: '1, c)', fallbackSize: 11 }]);
  });

  it('excludes raised author association markers from author-name size validation', () => {
    const nameRun = {
      text: 'Rachmad Andri Atmoko',
      fontName: 'Times New Roman',
      fontSize: 14,
      bold: false,
      italic: false,
      allCaps: false,
      verticalAlign: null,
      baselinePositionPt: null,
      sources: {}
    };
    const positionMarker = {
      text: '1, 2, a)',
      fontName: 'Times New Roman',
      fontSize: 11,
      bold: false,
      italic: false,
      allCaps: false,
      verticalAlign: null,
      baselinePositionPt: 5,
      sources: {}
    };

    const passing = validateRule('AUTHOR_NAME_SIZE_001', snapshot([
      paragraph({
        classification: 'authorName',
        text: 'Rachmad Andri Atmoko1, 2, a)',
        fontSize: null,
        effectiveRuns: [nameRun, positionMarker]
      })
    ]));
    expect(passing.status).toBe('passed');

    const failing = validateRule('AUTHOR_NAME_SIZE_001', snapshot([
      paragraph({
        classification: 'authorName',
        text: 'Rachmad Andri Atmoko1, 2, a)',
        fontSize: null,
        effectiveRuns: [{ ...nameRun, fontSize: 11 }, positionMarker]
      })
    ]));
    expect(failing.status).toBe('failed');
  });

  it('keeps figure-caption spacing validation while enforcing front-matter spacing separately', () => {
    const figure = paragraph({ classification: 'figureCaption', spaceBefore: 6 });
    expect(validateRule('FIGURE_CAPTION_SPACE_BEFORE_001', snapshot([figure])).status).toBe('passed');
    expect(strategy.rules.some(rule => rule.id === 'ABSTRACT_SPACE_BEFORE_001')).toBe(false);
    expect(strategy.rules.some(rule => rule.id === 'ABSTRACT_SPACE_AFTER_001')).toBe(false);
    expect(strategy.rules.some(rule => rule.id === 'AUTHOR_NAME_SPACING_001')).toBe(true);
    expect(strategy.rules.some(rule => rule.id === 'AUTHOR_AFFILIATION_SPACING_001')).toBe(true);
    expect(strategy.rules.some(rule => rule.id === 'AUTHOR_EMAIL_SPACING_001')).toBe(true);
    expect(strategy.rules.some(rule => rule.id === 'ABSTRACT_SPACING_001')).toBe(true);
    expect(strategy.rules.some(rule => rule.id === 'TITLE_SPACING_001')).toBe(false);
    expect(validateRule('FIGURE_CAPTION_SPACE_AFTER_001', snapshot([paragraph({ classification: 'figureCaption', spaceAfter: 0 })])).status).toBe('passed');
    expect(validateRule('TABLE_CAPTION_SPACE_BEFORE_001', snapshot([paragraph({ classification: 'tableCaption', spaceBefore: 6 })])).status).toBe('passed');
    expect(validateRule('TABLE_CAPTION_SPACE_AFTER_001', snapshot([paragraph({ classification: 'tableCaption', spaceAfter: 0 })])).status).toBe('passed');
    expect(validateRule('TITLE_SPACE_BEFORE_001', snapshot([paragraph({ classification: 'title', spaceBefore: 0 })])).status).toBe('passed');
    const titleSpacingFailure = validateRule('TITLE_SPACE_BEFORE_001', snapshot([paragraph({ classification: 'title', spaceBefore: 60 })]));
    expect(titleSpacingFailure.status).toBe('failed');
    expect(titleSpacingFailure.autoFixable).toBe(true);
    expect(strategy.rules).toHaveLength(174);
  });

  it('enforces 18 pt before and after spacing for author names and abstracts', () => {
    expect(validateRule('AUTHOR_NAME_SPACING_001', snapshot([
      paragraph({ classification: 'authorName', spaceBefore: 18, spaceAfter: 18 })
    ])).status).toBe('passed');
    const authorFailure = validateRule('AUTHOR_NAME_SPACING_001', snapshot([
      paragraph({ index: 4, classification: 'authorName', text: 'Author', spaceBefore: 0, spaceAfter: 12 })
    ]));
    expect(authorFailure.status).toBe('failed');
    expect(authorFailure.autoFixable).toBe(true);
    expect(authorFailure.paragraphIndices).toEqual([4]);

    expect(validateRule('ABSTRACT_SPACING_001', snapshot([
      paragraph({ classification: 'abstract', spaceBefore: 18, spaceAfter: 18 })
    ])).status).toBe('passed');
    expect(validateRule('ABSTRACT_SPACING_001', snapshot([
      paragraph({ classification: 'abstract', spaceBefore: 18, spaceAfter: 0 })
    ])).status).toBe('failed');

    expect(validateRule('AUTHOR_AFFILIATION_SPACING_001', snapshot([
      paragraph({ classification: 'authorAffiliation', spaceBefore: 0, spaceAfter: 0 })
    ])).status).toBe('passed');
    const affiliationFailure = validateRule('AUTHOR_AFFILIATION_SPACING_001', snapshot([
      paragraph({ index: 7, classification: 'authorAffiliation', spaceBefore: 12, spaceAfter: 6 })
    ]));
    expect(affiliationFailure.status).toBe('failed');
    expect(affiliationFailure.autoFixable).toBe(true);
    expect(affiliationFailure.assessmentRole).toBe('consistency');
    expect(affiliationFailure.message).toContain('0 pt before and after spacing');

    expect(validateRule('AUTHOR_EMAIL_SPACING_001', snapshot([
      paragraph({ classification: 'authorEmail', spaceBefore: 0, spaceAfter: 0 })
    ])).status).toBe('passed');
    const emailFailure = validateRule('AUTHOR_EMAIL_SPACING_001', snapshot([
      paragraph({ index: 8, classification: 'authorEmail', spaceBefore: 6, spaceAfter: 12 })
    ]));
    expect(emailFailure.status).toBe('failed');
    expect(emailFailure.autoFixable).toBe(true);
    expect(emailFailure.assessmentRole).toBe('consistency');
    expect(emailFailure.message).toContain('0 pt before and after spacing');
  });

  it('applies the first-post-Abstract Heading 1 spacing exception without conflicting with the general rule', () => {
    const document = snapshot([
      paragraph({ index: 0, text: 'Abstract. Summary', classification: 'abstract', spaceBefore: 18, spaceAfter: 18 }),
      paragraph({ index: 1, text: 'INTRODUCTION', classification: 'heading1', spaceBefore: 12, spaceAfter: 12 }),
      paragraph({ index: 2, text: 'METHODS', classification: 'heading1', spaceBefore: 12, spaceAfter: 12 })
    ]);
    const firstFailure = validateRule('FIRST_HEADING1_SPACE_BEFORE_001', document);
    expect(firstFailure.status).toBe('failed');
    expect(firstFailure.autoFixable).toBe(true);
    expect(firstFailure.assessmentRole).toBe('consistency');
    expect(firstFailure.paragraphIndices).toEqual([1]);

    const generalPass = validateRule('HEADING1_SPACE_BEFORE_001', document);
    expect(generalPass.status).toBe('passed');
    expect(generalPass.paragraphIndices).toEqual([]);

    const firstCorrect = snapshot([
      paragraph({ index: 0, text: 'Abstract. Summary', classification: 'abstract', spaceBefore: 18, spaceAfter: 18 }),
      paragraph({ index: 1, text: 'INTRODUCTION', classification: 'heading1', spaceBefore: 0, spaceAfter: 12 }),
      paragraph({ index: 2, text: 'METHODS', classification: 'heading1', spaceBefore: 0, spaceAfter: 12 })
    ]);
    expect(validateRule('FIRST_HEADING1_SPACE_BEFORE_001', firstCorrect).status).toBe('passed');
    const laterFailure = validateRule('HEADING1_SPACE_BEFORE_001', firstCorrect);
    expect(laterFailure.status).toBe('failed');
    expect(laterFailure.paragraphIndices).toEqual([2]);
  });

  it('keeps the first-heading exception out of AIP compliance scoring', () => {
    const document = snapshot([
      paragraph({ index: 0, text: 'Abstract. Summary', classification: 'abstract', spaceBefore: 18, spaceAfter: 18 }),
      paragraph({ index: 1, text: 'INTRODUCTION', classification: 'heading1', spaceBefore: 12, spaceAfter: 12 }),
      paragraph({ index: 2, text: 'METHODS', classification: 'heading1', spaceBefore: 12, spaceAfter: 12 })
    ]);
    const specialRule = strategy.rules.find(rule => rule.id === 'FIRST_HEADING1_SPACE_BEFORE_001');
    const complianceRule = strategy.rules.find(rule => rule.id === 'HEADING1_SPACE_AFTER_001');
    if (!specialRule || !complianceRule) throw new Error('Heading spacing score regression rules are missing.');
    const summary = ValidationEngine.validate({ ...strategy, rules: [specialRule, complianceRule] }, document);
    expect(summary.complianceScore).toBe(100);
    expect(summary.overallScore).toBe(50);
    expect(summary.unresolvedFailures).toBe(1);
    expect(summary.fixableFailures).toBe(1);
  });

  it('does not guess the first Heading 1 exception without an Abstract', () => {
    const finding = validateRule('FIRST_HEADING1_SPACE_BEFORE_001', snapshot([
      paragraph({ index: 0, text: 'INTRODUCTION', classification: 'heading1', spaceBefore: 12, spaceAfter: 12 }),
      paragraph({ index: 1, text: 'METHODS', classification: 'heading1', spaceBefore: 12, spaceAfter: 12 })
    ]));
    expect(finding.status).toBe('not_applicable');
  });

  it('targets the Heading 1 after Abstract rather than a Heading 1 before it', () => {
    const finding = validateRule('FIRST_HEADING1_SPACE_BEFORE_001', snapshot([
      paragraph({ index: 0, text: 'TEMPLATE HEADING', classification: 'heading1', spaceBefore: 12, spaceAfter: 12 }),
      paragraph({ index: 1, text: 'Abstract. Summary', classification: 'abstract', spaceBefore: 18, spaceAfter: 18 }),
      paragraph({ index: 2, text: 'INTRODUCTION', classification: 'heading1', spaceBefore: 12, spaceAfter: 12 })
    ]));
    expect(finding.status).toBe('failed');
    expect(finding.paragraphIndices).toEqual([2]);
  });

  it('changes only space-before for the first post-Abstract Heading 1', async () => {
    const finding = validateRule('FIRST_HEADING1_SPACE_BEFORE_001', snapshot([
      paragraph({ index: 1, text: 'Abstract. Summary', classification: 'abstract', spaceBefore: 18, spaceAfter: 18 }),
      paragraph({ index: 2, text: 'INTRODUCTION', classification: 'heading1', spaceBefore: 12, spaceAfter: 12, ooxmlAligned: true })
    ]));
    const applyBefore = vi.spyOn(WordDocumentAdapter, 'applyParagraphSpacingBefore').mockResolvedValue(undefined);
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot')
      .mockResolvedValueOnce(snapshot([
        paragraph({ index: 1, text: 'Abstract. Summary', classification: 'abstract', spaceBefore: 18, spaceAfter: 18 }),
        paragraph({ index: 2, text: 'INTRODUCTION', classification: 'heading1', spaceBefore: 12, spaceAfter: 12, ooxmlAligned: true })
      ]))
      .mockResolvedValueOnce(snapshot([
        paragraph({ index: 1, text: 'Abstract. Summary', classification: 'abstract', spaceBefore: 18, spaceAfter: 18 }),
        paragraph({ index: 2, text: 'INTRODUCTION', classification: 'heading1', spaceBefore: 0, spaceAfter: 12, ooxmlAligned: true })
      ]));
    try {
      const result = await CorrectionService.apply([finding]);
      expect(applyBefore).toHaveBeenCalledWith([2], 0);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
    } finally {
      snapshotSpy.mockRestore();
      applyBefore.mockRestore();
    }
  });

  it('refuses a stale first-post-Abstract Heading 1 spacing correction', async () => {
    const finding = validateRule('FIRST_HEADING1_SPACE_BEFORE_001', snapshot([
      paragraph({ index: 0, text: 'Abstract. Summary', classification: 'abstract', spaceBefore: 18, spaceAfter: 18 }),
      paragraph({ index: 1, text: 'INTRODUCTION', classification: 'heading1', spaceBefore: 12, spaceAfter: 12, ooxmlAligned: true })
    ]));
    const applyBefore = vi.spyOn(WordDocumentAdapter, 'applyParagraphSpacingBefore').mockResolvedValue(undefined);
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(snapshot([
      paragraph({ index: 0, text: 'Abstract. Summary', classification: 'abstract', spaceBefore: 18, spaceAfter: 18 }),
      paragraph({ index: 1, text: 'METHODS', classification: 'heading1', spaceBefore: 12, spaceAfter: 12, ooxmlAligned: true })
    ]));
    try {
      const result = await CorrectionService.apply([finding]);
      expect(applyBefore).not.toHaveBeenCalled();
      expect(result.executionFailures).toBe(1);
    } finally {
      snapshotSpy.mockRestore();
      applyBefore.mockRestore();
    }
  });

  it('enforces exactly one safe blank paragraph between affiliations and emails', () => {
    const safeBlank = (index: number) => paragraph({
      index,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const affiliation = paragraph({ index: 0, text: 'Affiliation', classification: 'authorAffiliation' });
    const email = paragraph({ index: 2, text: 'email@example.com', classification: 'authorEmail' });

    const missing = validateRule('AFFILIATION_EMAIL_BLANK_001', snapshot([affiliation, email]));
    expect(missing.status).toBe('failed');
    expect(missing.autoFixable).toBe(true);
    expect(analyzeAffiliationEmailBlankLine([affiliation, email]).state).toBe('missing');

    const exact = validateRule('AFFILIATION_EMAIL_BLANK_001', snapshot([affiliation, safeBlank(1), email]));
    expect(exact.status).toBe('passed');
    expect(analyzeAffiliationEmailBlankLine([affiliation, safeBlank(1), email]).state).toBe('passed');

    const extra = validateRule('AFFILIATION_EMAIL_BLANK_001', snapshot([affiliation, safeBlank(1), safeBlank(2), email]));
    expect(extra.status).toBe('failed');
    expect(extra.detected).toContain('2 safe empty paragraphs');
    expect(analyzeAffiliationEmailBlankLine([affiliation, safeBlank(1), safeBlank(2), email]).removableBlankParagraphIndices).toEqual([2]);
  });

  it('does not edit a structural blank separator between affiliation and email', () => {
    const structuralBlank = paragraph({
      index: 1,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: true,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: false
      }
    });
    const finding = validateRule('AFFILIATION_EMAIL_BLANK_001', snapshot([
      paragraph({ index: 0, text: 'Affiliation', classification: 'authorAffiliation' }),
      structuralBlank,
      paragraph({ index: 2, text: 'email@example.com', classification: 'authorEmail' })
    ]));
    expect(finding.status).toBe('not_evaluated');
    expect(finding.autoFixable).toBe(false);
  });

  it('accepts one leading manual break as the affiliation-email separator', () => {
    const affiliation = paragraph({ index: 0, text: 'Affiliation', classification: 'authorAffiliation' });
    const email = paragraph({
      index: 1,
      text: 'Corresponding author: email@example.com',
      classification: 'authorEmail',
      leadingManualLineBreakCount: 1,
      ooxmlAligned: true
    });
    const analysis = analyzeAffiliationEmailBlankLine([affiliation, email]);
    expect(analysis.state).toBe('passed');
    expect(analysis.separatorKind).toBe('leadingManualBreak');
    expect(validateRule('AFFILIATION_EMAIL_BLANK_001', snapshot([affiliation, email]))).toMatchObject({
      status: 'passed',
      detected: '1 leading manual line break — compliant'
    });
  });

  it('removes only a redundant safe blank when a leading break already exists', () => {
    const blank = paragraph({
      index: 1,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const analysis = analyzeAffiliationEmailBlankLine([
      paragraph({ index: 0, text: 'Affiliation', classification: 'authorAffiliation' }),
      blank,
      paragraph({ index: 2, text: 'Corresponding author: email@example.com', classification: 'authorEmail', leadingManualLineBreakCount: 1 })
    ]);
    expect(analysis.state).toBe('extra');
    expect(analysis.separatorKind).toBe('mixed');
    expect(analysis.removableBlankParagraphIndices).toEqual([1]);
  });

  it('blocks multiple leading breaks instead of guessing how to normalize them', () => {
    const analysis = analyzeAffiliationEmailBlankLine([
      paragraph({ index: 0, text: 'Affiliation', classification: 'authorAffiliation' }),
      paragraph({ index: 1, text: 'Corresponding author: email@example.com', classification: 'authorEmail', leadingManualLineBreakCount: 2 })
    ]);
    expect(analysis.state).toBe('blocked');
    expect(analysis.separatorKind).toBe('unknown');
    expect(analysis.blockedParagraphIndices).toEqual([1]);
  });

  it('counts only leading ordinary breaks in non-empty paragraph OOXML', () => {
    expect(leadingManualLineBreakCount('<w:p><w:r><w:br/><w:t>Corresponding author:</w:t></w:r></w:p>')).toBe(1);
    expect(leadingManualLineBreakCount('<w:p><w:r><w:br w:type="textWrapping"/><w:t>Corresponding author:</w:t></w:r></w:p>')).toBe(1);
    expect(leadingManualLineBreakCount('<w:p><w:r><w:t>Corresponding author:</w:t><w:br/><w:t>email@example.com</w:t></w:r></w:p>')).toBe(0);
    expect(leadingManualLineBreakCount('<w:p><w:r><w:br w:type="page"/><w:t>Corresponding author:</w:t></w:r></w:p>')).toBeNull();
    expect(leadingManualLineBreakCount('<w:p><w:r><w:br w:type="column"/><w:t>Corresponding author:</w:t></w:r></w:p>')).toBeNull();
    expect(leadingManualLineBreakCount('<w:p><w:r><w:br w:type="unexpected"/><w:t>Corresponding author:</w:t></w:r></w:p>')).toBeNull();
    expect(leadingManualLineBreakCount('<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:t>Corresponding author:</w:t></w:r></w:p>')).toBeNull();
    expect(leadingManualLineBreakCount('<w:p><w:r><w:br/><w:fldSimple w:instr="DATE"><w:r><w:t>October 2026</w:t></w:r></w:fldSimple><w:t>Corresponding author:</w:t></w:r></w:p>')).toBeNull();
  });

  it('uses the last affiliation and first following email for multiple front-matter lines', () => {
    const safeBlank = paragraph({
      index: 2,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const analysis = analyzeAffiliationEmailBlankLine([
      paragraph({ index: 0, text: 'Affiliation 1', classification: 'authorAffiliation' }),
      paragraph({ index: 1, text: 'Affiliation 2', classification: 'authorAffiliation' }),
      safeBlank,
      paragraph({ index: 3, text: 'one@example.com', classification: 'authorEmail' }),
      paragraph({ index: 4, text: 'two@example.com', classification: 'authorEmail' })
    ]);
    expect(analysis.state).toBe('passed');
    expect(analysis.lastAffiliationParagraphIndex).toBe(1);
    expect(analysis.firstEmailParagraphIndex).toBe(3);
  });

  it('enforces zero blank paragraphs across front-matter transitions without double-reporting affiliation to email', () => {
    const safeBlank = (index: number) => paragraph({
      index,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const paragraphs = [
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      safeBlank(1),
      paragraph({ index: 2, text: 'Author One', classification: 'authorName' }),
      safeBlank(3),
      paragraph({ index: 4, text: 'Author Two', classification: 'authorName' }),
      safeBlank(5),
      paragraph({ index: 6, text: 'Affiliation One', classification: 'authorAffiliation' }),
      paragraph({ index: 7, text: 'Affiliation Two', classification: 'authorAffiliation' }),
      safeBlank(8),
      safeBlank(9),
      paragraph({ index: 10, text: 'one@example.com', classification: 'authorEmail' }),
      safeBlank(11),
      paragraph({ index: 12, text: 'two@example.com', classification: 'authorEmail' }),
      safeBlank(13),
      paragraph({ index: 14, text: 'Abstract. Summary', classification: 'abstract' })
    ];
    const transitions = analyzeFrontMatterTransitions(paragraphs);
    const affiliationEmail = transitions.find(transition =>
      transition.fromType === 'authorAffiliation' && transition.toType === 'authorEmail'
    );
    expect(affiliationEmail?.expectedBlankCount).toBe(1);
    expect(affiliationEmail?.blankIndices).toEqual([8, 9]);

    const extra = validateRule('FRONTMATTER_EXTRA_BLANK_001', snapshot(paragraphs));
    expect(extra.status).toBe('failed');
    expect(extra.autoFixable).toBe(true);
    expect(extra.fixParagraphIndices).toEqual([1, 3, 5, 11, 13]);
    expect(extra.frontMatterFixTargets).toHaveLength(5);
    expect(extra.frontMatterFixTargets?.[0]).toMatchObject({
      fromType: 'title',
      toType: 'authorName',
      fromIndex: 0,
      toIndex: 2,
      blankIndices: [1],
      expectedBlankCount: 0
    });
    expect(extra.detected).not.toContain('authorAffiliation → authorEmail');

    const separator = validateRule('AFFILIATION_EMAIL_BLANK_001', snapshot(paragraphs));
    expect(separator.status).toBe('failed');
    expect(separator.detected).toContain('2 safe empty paragraphs');
  });

  it('passes the front-matter extra-blank rule when zero-blank transitions comply', () => {
    const compliant = snapshot([
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      paragraph({ index: 1, text: 'Author One', classification: 'authorName' }),
      paragraph({ index: 2, text: 'Affiliation', classification: 'authorAffiliation' }),
      paragraph({ index: 3, text: '', classification: 'other', isBlank: true, blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      } }),
      paragraph({ index: 4, text: 'email@example.com', classification: 'authorEmail' }),
      paragraph({ index: 5, text: 'Abstract. Summary', classification: 'abstract' })
    ]);
    expect(validateRule('FRONTMATTER_EXTRA_BLANK_001', compliant).status).toBe('passed');
    expect(validateRule('AFFILIATION_EMAIL_BLANK_001', compliant).status).toBe('passed');
  });

  it('does not auto-delete structural front-matter blanks', () => {
    const structuralBlank = paragraph({
      index: 1,
      text: '',
      classification: 'other',
      isBlank: true,
      ooxmlAligned: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: true,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: false
      }
    });
    const finding = validateRule('FRONTMATTER_EXTRA_BLANK_001', snapshot([
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      structuralBlank,
      paragraph({ index: 2, text: 'Author One', classification: 'authorName' })
    ]));
    expect(finding.status).toBe('not_evaluated');
    expect(finding.autoFixable).toBe(false);
  });

  it('does not treat unreliable front-matter blanks as safe fixes', () => {
    const unreliableBlank = paragraph({
      index: 1,
      text: '',
      classification: 'other',
      isBlank: true,
      ooxmlAligned: false,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const finding = validateRule('FRONTMATTER_EXTRA_BLANK_001', snapshot([
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      unreliableBlank,
      paragraph({ index: 2, text: 'Author One', classification: 'authorName' })
    ]));
    expect(finding.status).toBe('not_evaluated');
    expect(finding.autoFixable).toBe(false);
  });

  it('ignores non-blank content between front-matter components', () => {
    const finding = validateRule('FRONTMATTER_EXTRA_BLANK_001', snapshot([
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      paragraph({ index: 1, text: 'Unexpected note', classification: 'other' }),
      paragraph({ index: 2, text: 'Author One', classification: 'authorName' })
    ]));
    expect(finding.status).toBe('not_applicable');
  });

  it('separates AIP compliance from overall formatting when consistency fails', () => {
    const blank = paragraph({
      index: 1,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const author = paragraph({
      index: 2,
      text: 'Author One',
      classification: 'authorName',
      spaceBefore: 18,
      spaceAfter: 18
    });
    const doc = snapshot([
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      blank,
      author
    ]);
    const complianceRule = strategy.rules.find(rule => rule.id === 'AUTHOR_NAME_SPACING_001');
    const consistencyRule = strategy.rules.find(rule => rule.id === 'FRONTMATTER_EXTRA_BLANK_001');
    if (!complianceRule || !consistencyRule) throw new Error('Required score regression rules are missing.');

    const summary = ValidationEngine.validate({ ...strategy, rules: [complianceRule, consistencyRule] }, doc);
    expect(summary.complianceScore).toBe(100);
    expect(summary.score).toBe(100);
    expect(summary.overallScore).toBe(75);
    expect(summary.compliancePassed).toBe(1);
    expect(summary.complianceFailed).toBe(0);
    expect(summary.unresolvedFailures).toBe(1);
    expect(summary.fixableFailures).toBe(1);
  });

  it('excludes not-applicable and not-evaluated findings from deterministic score denominators', () => {
    const absent = ValidationEngine.validate({
      ...strategy,
      rules: [strategy.rules.find(rule => rule.id === 'FRONTMATTER_EXTRA_BLANK_001')!]
    }, snapshot([]));
    expect(absent.complianceScore).toBe(0);
    expect(absent.overallScore).toBe(0);
    expect(absent.unresolvedFailures).toBe(0);

    const structuralBlank = paragraph({
      index: 1,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: true,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: false
      }
    });
    const notEvaluated = ValidationEngine.validate({
      ...strategy,
      rules: [strategy.rules.find(rule => rule.id === 'FRONTMATTER_EXTRA_BLANK_001')!]
    }, snapshot([
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      structuralBlank,
      paragraph({ index: 2, text: 'Author One', classification: 'authorName' })
    ]));
    expect(notEvaluated.notEvaluated).toBe(1);
    expect(notEvaluated.complianceScore).toBe(0);
    expect(notEvaluated.overallScore).toBe(0);
  });

  it('classifies real front matter before applying the three front-matter rules', () => {
    const raw = (overrides: Partial<ParagraphSnapshot>) => {
      const { classification: _classification, ...withoutClassification } = paragraph(overrides);
      return withoutClassification;
    };
    const safeBlank = (index: number) => raw({
      index,
      text: '',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const classifyFrontMatter = (before: number, after: number, includeBlank: boolean) => {
      const input = [
        raw({ index: 0, text: 'Paper Title' }),
        raw({ index: 1, text: 'Author One', spaceBefore: before, spaceAfter: after }),
        raw({ index: 2, text: 'University of Example' }),
        ...(includeBlank ? [safeBlank(3)] : []),
        raw({ index: includeBlank ? 4 : 3, text: 'author@example.com' }),
        raw({ index: includeBlank ? 5 : 4, text: 'Abstract. Summary text', spaceBefore: before, spaceAfter: after })
      ];
      return snapshot(WordDocumentAdapter.classify(input));
    };

    const broken = classifyFrontMatter(0, 0, false);
    expect(broken.paragraphs.find(p => p.text === 'Author One')?.classification).toBe('authorName');
    expect(broken.paragraphs.find(p => p.text === 'University of Example')?.classification).toBe('authorAffiliation');
    expect(broken.paragraphs.find(p => p.text === 'author@example.com')?.classification).toBe('authorEmail');
    expect(broken.paragraphs.find(p => p.text.startsWith('Abstract.'))?.classification).toBe('abstract');
    expect(validateRule('AUTHOR_NAME_SPACING_001', broken).status).toBe('failed');
    expect(validateRule('AUTHOR_NAME_SPACING_001', broken).autoFixable).toBe(true);
    expect(validateRule('AFFILIATION_EMAIL_BLANK_001', broken).status).toBe('failed');
    expect(validateRule('AFFILIATION_EMAIL_BLANK_001', broken).autoFixable).toBe(true);
    expect(validateRule('ABSTRACT_SPACING_001', broken).status).toBe('failed');
    expect(validateRule('ABSTRACT_SPACING_001', broken).autoFixable).toBe(true);

    const compliant = classifyFrontMatter(18, 18, true);
    expect(validateRule('AUTHOR_NAME_SPACING_001', compliant).status).toBe('passed');
    expect(validateRule('AFFILIATION_EMAIL_BLANK_001', compliant).status).toBe('passed');
    expect(validateRule('ABSTRACT_SPACING_001', compliant).status).toBe('passed');
  });

  it('recognizes normalized section aliases and combined section families', () => {
    expect(normalizeSectionHeading('1. Results and Discussion:')).toBe('RESULTS AND DISCUSSION');
    expect(normalizeSectionHeading('1 Introduction')).toBe('INTRODUCTION');
    expect(normalizeSectionHeading('1.2 Background')).toBe('BACKGROUND');
    expect(normalizeSectionHeading('2.3.1 Methodology')).toBe('METHODOLOGY');
    expect(normalizeSectionHeading('II Results')).toBe('RESULTS');
    expect(normalizeSectionHeading('IC Results')).toBe('IC RESULTS');
    expect(canonicalSectionsForHeading('1. Results and Discussion')).toEqual([
      'results_discussion',
      'results',
      'discussion'
    ]);
    expect(canonicalSectionsForHeading('Conclusion and Future Work')).toEqual([
      'conclusion_future_work',
      'conclusion',
      'future_work'
    ]);
  });

  it('recognizes first-level section metadata without imposing a fixed IMRaD whitelist', () => {
    const classified = WordDocumentAdapter.classify([
      paragraph({ index: 0, text: 'Title' }),
      paragraph({ index: 1, text: 'INTRODUCTION', style: 'Heading 1', styleBuiltIn: 'Heading 1', allCaps: true }),
      paragraph({ index: 2, text: 'MODEL FORMULATION', style: 'Heading 1', styleBuiltIn: 'Heading 1', allCaps: true }),
      paragraph({ index: 3, text: 'RESULTS AND DISCUSSION', style: 'Heading 1', styleBuiltIn: 'Heading 1', allCaps: true })
    ]);

    expect(classified[1].classification).toBe('heading1');
    expect(classified[1].sectionFamilies).toContain('introduction');
    expect(classified[2].classification).toBe('heading1');
    expect(classified[2].sectionFamilies).toEqual([]);
    expect(classified[2].sectionConfidence).toBe('high');
    expect(classified[3].sectionFamilies).toEqual(['results_discussion', 'results', 'discussion']);
  });

  it('requires one authoritative Abstract presence rule', () => {
    expect(strategy.rules.some(rule => rule.id === 'ABSTRACT_REQUIRED_001')).toBe(false);
    expect(strategy.rules.filter(rule => rule.property === 'document.abstractPresence')).toHaveLength(1);

    const missing = validateRule('ABSTRACT_PRESENCE_001', snapshot([
      paragraph({ index: 0, text: 'Title', classification: 'title' }),
      paragraph({ index: 1, text: 'MODEL FORMULATION', classification: 'heading1', sectionConfidence: 'high' })
    ]));
    expect(missing.status).toBe('failed');
    expect(missing.autoFixable).toBe(false);

    const present = validateRule('ABSTRACT_PRESENCE_001', snapshot([
      paragraph({ index: 0, text: 'Title', classification: 'title' }),
      paragraph({ index: 1, text: 'Abstract. Summary', classification: 'abstract' })
    ]));
    expect(present.status).toBe('passed');
  });

  it('detects required structural section families without requiring literal headings', () => {
    const section = (index: number, text: string) => paragraph({
      index,
      text,
      classification: 'heading1',
      sectionFamilies: canonicalSectionsForHeading(text),
      sectionConfidence: 'high'
    });
    const positiveCases: Array<{ ruleId: string; headings: Array<[string, string[]]> }> = [
      {
        ruleId: 'INTRODUCTION_PRESENCE_001',
        headings: [
          ['INTRODUCTION', ['introduction']],
          ['GENERAL INTRODUCTION', ['introduction']],
          ['INTRODUCTORY REMARKS', ['introduction']],
          ['OVERVIEW', ['introduction']]
        ]
      },
      {
        ruleId: 'METHOD_SECTION_PRESENCE_001',
        headings: [
          ['METHODOLOGY', ['methodology']],
          ['MATERIALS AND METHODS', ['materials_methods']],
          ['EXPERIMENTAL METHOD', ['experimental_method']],
          ['PROPOSED SYSTEM', ['proposed_solution']],
          ['SYSTEM DESIGN', ['system_architecture']],
          ['IMPLEMENTATION', ['implementation']]
        ]
      },
      {
        ruleId: 'RESULTS_SECTION_PRESENCE_001',
        headings: [
          ['RESULTS', ['results']],
          ['RESULTS AND DISCUSSION', ['results_discussion', 'results', 'discussion']],
          ['PERFORMANCE EVALUATION', ['evaluation']],
          ['ANALYSIS', ['analysis']]
        ]
      },
      {
        ruleId: 'CONCLUSION_PRESENCE_001',
        headings: [
          ['CONCLUSION', ['conclusion']],
          ['CONCLUSIONS', ['conclusion']],
          ['SUMMARY', ['conclusion']],
          ['SUMMARY AND CONCLUSION', ['conclusion']],
          ['SUMMARY AND CONCLUSIONS', ['conclusion']],
          ['CONCLUSION AND FUTURE WORK', ['conclusion_future_work', 'conclusion', 'future_work']]
        ]
      }
    ];

    let index = 1;
    for (const testCase of positiveCases) {
      for (const [text, families] of testCase.headings) {
        expect(canonicalSectionsForHeading(text), `${testCase.ruleId}: ${text}`).toEqual(families);
        const finding = validateRule(testCase.ruleId, snapshot([section(index, text)]));
        expect(finding.status, `${testCase.ruleId}: ${text}`).toBe('passed');
        expect(finding.assessmentRole).toBe('consistency');
        expect(finding.paragraphIndices).toEqual([index]);
        index += 1;
      }
    }
  });

  it('does not let Heading 2 or Heading 3 satisfy required major sections', () => {
    const section = (index: number, text: string, classification: 'heading1' | 'heading2' | 'heading3') => paragraph({
      index,
      text,
      classification,
      sectionFamilies: canonicalSectionsForHeading(text),
      sectionConfidence: 'high'
    });
    const nestedCases: Array<[string, string, 'heading2' | 'heading3']> = [
      ['INTRODUCTION_PRESENCE_001', 'INTRODUCTION', 'heading2'],
      ['METHOD_SECTION_PRESENCE_001', 'METHODS', 'heading2'],
      ['RESULTS_SECTION_PRESENCE_001', 'RESULTS', 'heading3'],
      ['CONCLUSION_PRESENCE_001', 'CONCLUSION', 'heading2']
    ];

    for (const [ruleId, text, classification] of nestedCases) {
      expect(validateRule(ruleId, snapshot([section(1, text, classification)])).status)
        .toBe('failed');
    }

    expect(validateRule(
      'INTRODUCTION_PRESENCE_001',
      snapshot([section(1, 'INTRODUCTION', 'heading1')])
    ).status).toBe('passed');
  });

  it('reports missing structural groups without accepting unrelated section families', () => {
    const section = (text: string) => paragraph({
      index: 1,
      text,
      classification: 'heading1',
      sectionFamilies: canonicalSectionsForHeading(text),
      sectionConfidence: 'high'
    });
    const negativeCases: Array<[string, string, string[]]> = [
      ['INTRODUCTION_PRESENCE_001', 'RELATED WORK', ['related_work']],
      ['INTRODUCTION_PRESENCE_001', 'BACKGROUND', ['background']],
      ['METHOD_SECTION_PRESENCE_001', 'BACKGROUND', ['background']],
      ['METHOD_SECTION_PRESENCE_001', 'RELATED WORK', ['related_work']],
      ['RESULTS_SECTION_PRESENCE_001', 'DISCUSSION', ['discussion']],
      ['RESULTS_SECTION_PRESENCE_001', 'FUTURE WORK', ['future_work']],
      ['CONCLUSION_PRESENCE_001', 'FUTURE WORK', ['future_work']],
      ['CONCLUSION_PRESENCE_001', 'DISCUSSION', ['discussion']]
    ];

    for (const [ruleId, text, families] of negativeCases) {
      expect(canonicalSectionsForHeading(text)).toEqual(families);
      const finding = validateRule(ruleId, snapshot([section(text)]));
      expect(finding.status, ruleId).toBe('failed');
      expect(finding.assessmentRole).toBe('consistency');
      expect(finding.paragraphIndices).toEqual([]);
      expect(finding.detected).toContain('No recognized');
    }
  });

  it('requires a References/Bibliography heading only when citations exist', () => {
    const noCitations = validateRule('REFERENCE_SECTION_CONSISTENCY_001', snapshot([
      paragraph({ index: 0, text: 'Title', classification: 'title' })
    ]));
    expect(noCitations.status).toBe('not_applicable');

    const withCitation = snapshot([
      paragraph({ index: 0, text: 'Title', classification: 'title' })
    ]);
    withCitation.crossReferences = {
      referenceDefinitions: [],
      referenceCitations: [{ kind: 'reference', numbers: [1], paragraphIndex: 0, rawText: '[1]' }],
      referenceCitationStyle: 'square_bracket_numeric',
      figureDefinitions: [],
      figureCitations: [],
      tableDefinitions: [],
      tableCitations: [],
      referenceCoverageMissing: [],
      invalidReferenceCitations: [],
      referenceSequenceViolations: [],
      figureCoverageMissing: [],
      tableCoverageMissing: [],
      figurePlacement: [],
      tablePlacement: [],
      figurePriorCitationFailures: [],
      figurePriorCitationUnknown: []
    };
    expect(validateRule('REFERENCE_SECTION_CONSISTENCY_001', withCitation).status).toBe('failed');
    withCitation.crossReferences!.invalidReferenceCitations = [withCitation.crossReferences!.referenceCitations[0]];
    expect(validateRule('REFERENCE_SECTION_CONSISTENCY_001', withCitation).message).toContain('1 in-text citation');

    const withBibliography = WordDocumentAdapter.classify([
      paragraph({ index: 0, text: 'Title' }),
      paragraph({ index: 1, text: 'Abstract. Summary' }),
      paragraph({ index: 2, text: 'BIBLIOGRAPHY', style: 'Heading 1', styleBuiltIn: 'Heading 1', allCaps: true }),
      paragraph({ index: 3, text: '1. Example reference', style: 'Normal' })
    ]);
    expect(withBibliography[2].classification).toBe('referencesHeading');
    expect(withBibliography[3].classification).toBe('reference');
  });

  it('treats the section structure review as advisory and supports combined headings', () => {
    const finding = validateRule('SECTION_STRUCTURE_REVIEW_001', snapshot([
      paragraph({
        index: 1,
        text: 'RESULTS AND DISCUSSION',
        classification: 'heading1',
        sectionFamilies: ['results_discussion', 'results', 'discussion'],
        sectionNormalizedTitle: 'RESULTS AND DISCUSSION',
        sectionConfidence: 'high'
      })
    ]));
    expect(finding.status).toBe('manual_review');
    expect(finding.assessmentRole).toBe('consistency');
    expect(finding.message).toContain('results');
    expect(finding.message).toContain('discussion');
  });

  it('does not count known section aliases at Heading 2/3 level as major sections', () => {
    const finding = validateRule('SECTION_STRUCTURE_REVIEW_001', snapshot([
      paragraph({
        index: 1,
        text: 'BACKGROUND',
        classification: 'heading2',
        sectionFamilies: ['background'],
        sectionConfidence: 'high'
      }),
      paragraph({
        index: 2,
        text: 'METHODOLOGY',
        classification: 'heading3',
        sectionFamilies: ['methodology'],
        sectionConfidence: 'high'
      })
    ]));
    expect(finding.status).toBe('not_applicable');
  });

  it('detects safe Heading 1 style candidates without promoting subsections or layout content', () => {
    const wrongStyle = validateRule('HEADING1_STYLE_001', snapshot([
      paragraph({
        index: 1,
        text: 'INTRODUCTION',
        classification: 'body',
        style: 'Normal',
        sectionFamilies: ['introduction'],
        sectionConfidence: 'high',
        heading1Confidence: 'high',
        ooxmlAligned: true
      })
    ]));
    expect(wrongStyle.status).toBe('failed');
    expect(wrongStyle.autoFixable).toBe(true);
    expect(wrongStyle.paragraphIndices).toEqual([1]);

    const correctStyle = validateRule('HEADING1_STYLE_001', snapshot([
      paragraph({
        index: 1,
        text: 'INTRODUCTION',
        classification: 'heading1',
        style: 'Heading 1',
        styleBuiltIn: 'Heading 1',
        sectionFamilies: ['introduction'],
        sectionConfidence: 'high',
        heading1Confidence: 'high'
      })
    ]));
    expect(correctStyle.status).toBe('passed');

    for (const classification of ['heading2', 'heading3'] as const) {
      expect(validateRule('HEADING1_STYLE_001', snapshot([
        paragraph({
          index: 2,
          text: 'BACKGROUND',
          classification,
          style: 'Normal',
          sectionFamilies: ['background'],
          sectionConfidence: 'high'
        })
      ])).status).toBe('not_applicable');
    }

    expect(validateRule('HEADING1_STYLE_001', snapshot([
      paragraph({
        index: 3,
        text: 'TABLE HEADER',
        classification: 'other',
        tableNestingLevel: 1,
        sectionFamilies: ['other'],
        sectionConfidence: 'high'
      })
    ])).status).toBe('not_applicable');
  });

  it('allows a custom first-level heading candidate without requiring an ontology alias', () => {
    const classified = WordDocumentAdapter.classify([
      paragraph({ index: 0, text: 'Paper title' }),
      paragraph({
        index: 1,
        text: 'RESOURCE-AWARE FEDERATED ARCHITECTURE',
        style: 'Normal',
        styleBuiltIn: 'Normal',
        alignment: 'center',
        fontSize: 12,
        bold: true,
        allCaps: true,
        ooxmlAligned: true
      })
    ]);
    expect(classified[1].classification).toBe('body');
    expect(classified[1].sectionFamilies).toEqual([]);
    expect(classified[1].heading1Confidence).toBe('high');
    const finding = validateRule('HEADING1_STYLE_001', snapshot(classified));
    expect(finding.status).toBe('failed');
    expect(finding.autoFixable).toBe(true);
  });

  it('does not auto-fix a weak ontology alias that lacks heading evidence', () => {
    const classified = WordDocumentAdapter.classify([
      paragraph({
        index: 0,
        text: 'Overview',
        style: 'Normal',
        styleBuiltIn: 'Normal',
        alignment: 'left',
        fontSize: 10,
        bold: false,
        allCaps: false,
        ooxmlAligned: true
      })
    ]);
    expect(classified[0].heading1Confidence).not.toBe('high');
    expect(validateRule('HEADING1_STYLE_001', snapshot(classified)).status).toBe('not_applicable');
  });

  it('never promotes titles or captions to Heading 1', () => {
    for (const classification of ['title', 'tableCaption', 'figureCaption'] as const) {
      const finding = validateRule('HEADING1_STYLE_001', snapshot([
        paragraph({
          index: 0,
          text: classification === 'title' ? 'RESOURCE-AWARE SECURITY IN FEDERATED LEARNING' : 'TABLE 1. RESULTS',
          classification,
          style: 'Normal',
          fontSize: 12,
          bold: true,
          alignment: 'center',
          allCaps: true,
          heading1Confidence: 'high',
          ooxmlAligned: true
        })
      ]));
      expect(finding.status).toBe('not_applicable');
      expect(finding.autoFixable).toBe(false);
    }
  });

  it('keeps a References heading eligible for Heading 1 style correction', () => {
    const finding = validateRule('HEADING1_STYLE_001', snapshot([
      paragraph({
        index: 1,
        text: 'REFERENCES',
        classification: 'referencesHeading',
        style: 'Normal',
        heading1Confidence: 'high',
        ooxmlAligned: true
      })
    ]));
    expect(finding.status).toBe('failed');
    expect(finding.autoFixable).toBe(true);
  });

  it('offers a conditional safe fix for parseable Roman table labels', () => {
    const roman = validateRule('TABLE_CAPTION_PREFIX_001', snapshot([
      paragraph({ classification: 'tableCaption', text: 'TABLE I. Methods', tablePrefixAllCaps: true })
    ]));
    expect(roman.status).toBe('failed');
    expect(roman.autoFixable).toBe(true);
    expect(roman.fixParagraphIndices).toEqual([0]);

    const fieldBacked = validateRule('TABLE_CAPTION_PREFIX_001', snapshot([
      paragraph({
        classification: 'tableCaption',
        text: 'TABLE I. Methods',
        tablePrefixAllCaps: true,
        fieldTextSpans: [{ start: 6, end: 7 }]
      })
    ]));
    expect(fieldBacked.status).toBe('failed');
    expect(fieldBacked.autoFixable).toBe(false);
  });

  it('does not offer safe fixes for non-canonical Roman table labels', () => {
    for (const roman of ['IIII', 'IC', 'VX', 'IIV']) {
      const finding = validateRule('TABLE_CAPTION_PREFIX_001', snapshot([
        paragraph({
          classification: 'tableCaption',
          text: `TABLE ${roman}. Results`
        })
      ]));

      expect(finding.status).toBe('failed');
      expect(finding.autoFixable).toBe(false);
    }
  });

  it('detects repeated prose spaces and caption leading whitespace conservatively', () => {
    const repeated = validateRule('TEXT_MULTIPLE_SPACES_001', snapshot([
      paragraph({ index: 0, text: 'This  is prose', classification: 'body' }),
      paragraph({ index: 1, text: 'A  table cell', classification: 'other', tableNestingLevel: 1 })
    ]));
    expect(repeated.status).toBe('failed');
    expect(repeated.paragraphIndices).toEqual([0]);
    expect(repeated.assessmentRole).toBe('consistency');
    expect(validateRule('FIGURE_CAPTION_LEADING_WHITESPACE_001', snapshot([
      paragraph({ text: '   FIGURE 1. Result', classification: 'figureCaption' })
    ])).status).toBe('failed');
    expect(validateRule('TABLE_CAPTION_LEADING_WHITESPACE_001', snapshot([
      paragraph({ text: 'TABLE 1. Result', classification: 'tableCaption' })
    ])).status).toBe('passed');
    expect(collapseRepeatedOrdinarySpaces('This  is  a  paragraph.')).toBe('This is a paragraph.');
    expect(collapseRepeatedOrdinarySpaces('A  B  C  D  E  F  G  H  I  J')).toBe('A B C D E F G H I J');
    expect(collapseRepeatedOrdinarySpaces('word   word')).toBe('word word');
    expect(collapseRepeatedOrdinarySpaces('word word')).toBe('word word');
  });

  it('detects only safely removable body-to-body blank paragraphs', () => {
    const safeBlank = (index: number, text = '') => paragraph({
      index,
      text,
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const body = (index: number, text: string) => paragraph({ index, text, classification: 'body' });

    const finding = validateRule('EXTRA_BLANK_PARAGRAPH_001', snapshot([
      body(0, 'Paragraph A'),
      safeBlank(1),
      safeBlank(2, '  '),
      body(3, 'Paragraph B')
    ]));
    expect(finding.status).toBe('failed');
    expect(finding.paragraphIndices).toEqual([1, 2]);
    expect(finding.autoFixable).toBe(true);
    expect(finding.assessmentRole).toBe('consistency');

    expect(findExtraBlankBodyParagraphIndices([
      body(0, 'Heading-like body'),
      safeBlank(1),
      paragraph({ index: 2, text: 'Heading', classification: 'heading1' })
    ])).toEqual([]);
    expect(findExtraBlankBodyParagraphIndices([
      body(0, 'Body'),
      safeBlank(1),
      paragraph({ index: 2, text: 'FIGURE 1. Image', classification: 'figureCaption' })
    ])).toEqual([]);
    expect(findExtraBlankBodyParagraphIndices([
      paragraph({ index: 0, text: 'Item A', classification: 'body', isListItem: true }),
      safeBlank(1),
      paragraph({ index: 2, text: 'Item B', classification: 'body', isListItem: true })
    ])).toEqual([]);
  });

  it('repairs only the target paragraph in a Flat OPC package', () => {
    const flatOpc = `
      <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/xml">
          <pkg:xmlData><w:styles><w:style><w:pPr><w:ind w:left="999"/></w:pPr></w:style></w:styles></pkg:xmlData>
        </pkg:part>
        <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/xml">
          <pkg:xmlData><w:document><w:body>
            <w:p><w:r><w:t>Heading</w:t></w:r></w:p>
            <w:p><w:pPr><w:ind w:left="720" w:hanging="360"/><w:numPr><w:ilvl w:val="0"/></w:numPr></w:pPr><w:hyperlink><w:r><w:t>Target reference</w:t></w:r></w:hyperlink></w:p>
          </w:body></w:document></pkg:xmlData>
        </pkg:part>
      </pkg:package>`;

    const repaired = repairParagraphIndentOoxml(flatOpc, 1, 'Target reference');
    expect(repaired.changed).toBe(true);
    expect(repaired.xml).toContain('Target reference');
    expect(repaired.xml).not.toContain('<w:ind');
    expect(repaired.xml).toContain('<w:numPr>');
    expect(repaired.xml).toContain('<w:hyperlink>');

    const repairedByText = repairParagraphIndentOoxml(flatOpc, 60, 'Target reference');
    expect(repairedByText.changed).toBe(true);
    expect(repairedByText.xml).toContain('Target reference');

    for (const originalIndex of [12, 60]) {
      const singleParagraphPackage = flatOpc.replace(
        /\s*<w:p><w:r><w:t>Heading<\/w:t><\/w:r><\/w:p>/,
        ''
      );
      const repairedSingle = repairParagraphIndentOoxml(singleParagraphPackage, originalIndex);
      expect(repairedSingle.changed).toBe(true);
      expect(repairedSingle.xml).toContain('Target reference');
    }
  });

  it('withholds indentation replacement for relationship-backed Flat OPC content', () => {
    const hyperlinkPackage = `
      <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/xml">
          <pkg:xmlData><w:document><w:body><w:p><w:pPr><w:ind w:left="720"/></w:pPr><w:hyperlink r:id="rId44"><w:r><w:t>Linked reference</w:t></w:r></w:hyperlink></w:p></w:body></w:document></pkg:xmlData>
        </pkg:part>
        <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
          <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId44" Target="https://example.com/doi" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" TargetMode="External"/></Relationships></pkg:xmlData>
        </pkg:part>
      </pkg:package>`;

    expect(() => repairParagraphIndentOoxml(hyperlinkPackage, 60, 'Linked reference'))
      .toThrow('relationship-backed content');

    const olePackage = hyperlinkPackage.replace(
      /<w:hyperlink[\s\S]*?<\/w:hyperlink>/,
      '<w:object r:id="rId44"/>'
    );
    expect(() => repairParagraphIndentOoxml(olePackage, 12, ''))
      .toThrow('relationship-backed content');

    const alternateContentPackage = hyperlinkPackage
      .replace(
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
      )
      .replace(
        /<w:hyperlink[\s\S]*?<\/w:hyperlink>/,
        '<mc:AlternateContent r:id="rId44"><mc:Choice Requires="x"><w:r><w:t>Linked reference</w:t></w:r></mc:Choice></mc:AlternateContent>'
      );
    expect(() => repairParagraphIndentOoxml(alternateContentPackage, 12, 'Linked reference'))
      .toThrow('relationship-backed content');
  });

  it('does not delete structural or visually empty non-paragraph content', () => {
    const whitespace = inspectBlankParagraph(
      '<w:p><w:r><w:t xml:space="preserve"> &#xA0; </w:t></w:r></w:p>',
      ' \u00A0 '
    );
    expect(whitespace.whitespaceOnly).toBe(true);
    expect(whitespace.safeToDelete).toBe(true);

    const pageBreak = inspectBlankParagraph('<w:p><w:r><w:br w:type="page"/></w:r></w:p>', '');
    expect(pageBreak.hasExplicitPageBreak).toBe(true);
    expect(pageBreak.hasLastRenderedPageBreak).toBe(false);
    expect(pageBreak.safeToDelete).toBe(false);

    const renderedBreak = inspectBlankParagraph(
      '<w:p><w:r><w:lastRenderedPageBreak/></w:r></w:p>',
      ''
    );
    expect(renderedBreak.hasExplicitPageBreak).toBe(false);
    expect(renderedBreak.hasLastRenderedPageBreak).toBe(true);
    expect(renderedBreak.safeToDelete).toBe(true);
    expect(validateRule('FIGURE_CAPTION_BODY_GAP_001', snapshot([
      paragraph({ index: 0, text: 'FIGURE 1. Result', classification: 'figureCaption' }),
      paragraph({
        index: 1,
        text: '',
        classification: 'other',
        isBlank: true,
        blankParagraph: renderedBreak
      }),
      paragraph({ index: 2, text: 'The result shows...', classification: 'body' })
    ])).status).toBe('passed');

    const manualBreak = inspectBlankParagraph('<w:p><w:r><w:br/></w:r></w:p>', '');
    expect(manualBreak.hasManualBreak).toBe(true);
    expect(manualBreak.safeToDelete).toBe(false);

    const unsafeBlank = paragraph({
      index: 1,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: { ...manualBreak, whitespaceOnly: true }
    });
    const safeBlank = paragraph({
      index: 2,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    expect(findExtraBlankBodyParagraphIndices([
      paragraph({ index: 0, text: 'Body', classification: 'body' }),
      safeBlank,
      unsafeBlank,
      paragraph({ index: 3, text: 'Next body', classification: 'body' })
    ])).toEqual([]);

    const body = paragraph({ index: 0, text: 'Body', classification: 'body', spaceAfter: 12 });
    const nextBody = paragraph({ index: 1, text: 'Next body', classification: 'body' });
    expect(findExtraBlankBodyParagraphIndices([body, nextBody])).toEqual([]);
  });

  it('deletes blank-paragraph targets from highest index to lowest', () => {
    expect(WordDocumentAdapter.sortParagraphIndicesForDeletion([10, 20, 10, 30]))
      .toEqual([30, 20, 10]);
  });

  it('requires exactly one safe blank paragraph between a figure caption and body text', async () => {
    const blank = (index: number) => paragraph({
      index,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const caption = paragraph({ index: 0, text: 'FIGURE 1. Result', classification: 'figureCaption' });
    const body = paragraph({ index: 2, text: 'The result shows...', classification: 'body' });

    const missing = validateRule('FIGURE_CAPTION_BODY_GAP_001', snapshot([caption, body]));
    expect(missing.status).toBe('failed');
    expect(missing.paragraphIndices).toEqual([0]);
    expect(missing.autoFixable).toBe(true);

    const exact = validateRule('FIGURE_CAPTION_BODY_GAP_001', snapshot([caption, blank(1), body]));
    expect(exact.status).toBe('passed');

    const extra = validateRule('FIGURE_CAPTION_BODY_GAP_001', snapshot([
      caption,
      blank(1),
      blank(2),
      paragraph({ index: 3, text: 'The result shows...', classification: 'body' })
    ]));
    expect(extra.status).toBe('failed');
    expect(extra.detected).toContain('extra blank paragraph');

    for (const classification of ['heading1', 'heading2', 'heading3'] as const) {
      expect(validateRule('FIGURE_CAPTION_BODY_GAP_001', snapshot([
        caption,
        paragraph({ index: 1, text: classification, classification })
      ])).status).toBe('not_applicable');
    }
    expect(validateRule('FIGURE_CAPTION_BODY_GAP_001', snapshot([
      caption,
      paragraph({ index: 1, text: 'TABLE 1. Results', classification: 'tableCaption' })
    ])).status).toBe('not_applicable');

    const pageBreak = paragraph({
      index: 1,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: true,
        hasExplicitPageBreak: true,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: false
      }
    });
    expect(validateRule('FIGURE_CAPTION_BODY_GAP_001', snapshot([
      caption,
      pageBreak,
      paragraph({ index: 2, text: 'The result shows...', classification: 'body' })
    ])).status).toBe('not_evaluated');

    expect(validateRule('FIGURE_CAPTION_BODY_GAP_001', snapshot([
      caption,
      paragraph({
        index: 1,
        text: 'The result starts on a new page.',
        classification: 'body',
        hasPageBreakBefore: true
      })
    ])).status).toBe('not_evaluated');

    const captionWithBreak = paragraph({
      ...caption,
      blankParagraph: {
        whitespaceOnly: false,
        hasManualBreak: false,
          hasExplicitPageBreak: true,
          hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: false
      }
    });
    expect(validateRule('FIGURE_CAPTION_BODY_GAP_001', snapshot([
      captionWithBreak,
      body
    ])).status).toBe('not_evaluated');

    const parsedPageBreakBefore = parseParagraphOoxml(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>',
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    );
    expect(parsedPageBreakBefore[0].hasPageBreakBefore).toBe(true);

    expect(findFigureCaptionBodyGapIssues([caption, blank(1), body])).toHaveLength(0);

    const source = await readFile('src/word/WordDocumentAdapter.ts', 'utf8');
    expect(source).toContain('applyFigureCaptionBodyGaps');
    expect(source).toContain("paragraph.insertParagraph('', Word.InsertLocation.after)");
    expect(source).toContain('Blank paragraph ${index + 1} is no longer available. Revalidate the document.');
  });

  it('requires zero safe blank paragraphs between an external table caption and table', () => {
    const blank = (index: number, safeToDelete = true) => paragraph({
      index,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete
      }
    });
    const caption = paragraph({ index: 0, text: 'TABLE 1. Results', classification: 'tableCaption' });

    const makeDocument = (paragraphs: ParagraphSnapshot[], blocks: NonNullable<DocumentSnapshot['layout']['documentBlocks']>) => {
      const doc = snapshot(paragraphs);
      doc.layout.documentBlocks = blocks;
      doc.crossReferences = analyzeCrossReferences(paragraphs, blocks);
      return doc;
    };

    const direct = makeDocument(
      [caption],
      [
        { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, text: caption.text },
        { blockIndex: 1, kind: 'table', tableIndex: 0 }
      ]
    );
    expect(validateRule('TABLE_CAPTION_TABLE_GAP_001', direct).status).toBe('passed');

    const oneBlank = makeDocument(
      [caption, blank(1)],
      [
        { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, text: caption.text },
        { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1 },
        { blockIndex: 2, kind: 'table', tableIndex: 0 }
      ]
    );
    const oneBlankFinding = validateRule('TABLE_CAPTION_TABLE_GAP_001', oneBlank);
    expect(oneBlankFinding.status).toBe('failed');
    expect(oneBlankFinding.autoFixable).toBe(true);
    expect(oneBlankFinding.paragraphIndices).toEqual([0]);
    expect(oneBlankFinding.detected).toContain('1 extra blank paragraph');

    const twoBlanks = makeDocument(
      [caption, blank(1), blank(2)],
      [
        { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, text: caption.text },
        { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1 },
        { blockIndex: 2, kind: 'paragraph', paragraphIndex: 2 },
        { blockIndex: 3, kind: 'table', tableIndex: 0 }
      ]
    );
    const twoBlankAnalysis = analyzeTableCaptionTableGaps(twoBlanks);
    expect(twoBlankAnalysis.issues[0].blankParagraphIndices).toEqual([1, 2]);
    expect(validateRule('TABLE_CAPTION_TABLE_GAP_001', twoBlanks).detected).toContain('2 extra blank paragraphs');

    const structural = makeDocument(
      [caption, blank(1, false)],
      [
        { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, text: caption.text },
        { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1 },
        { blockIndex: 2, kind: 'table', tableIndex: 0 }
      ]
    );
    expect(validateRule('TABLE_CAPTION_TABLE_GAP_001', structural).status).toBe('not_evaluated');

    const interveningText = makeDocument(
      [caption, paragraph({ index: 1, text: 'Explanatory text', classification: 'body' })],
      [
        { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, text: caption.text },
        { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1, text: 'Explanatory text' },
        { blockIndex: 2, kind: 'table', tableIndex: 0 }
      ]
    );
    expect(validateRule('TABLE_CAPTION_TABLE_GAP_001', interveningText).status).toBe('not_applicable');

    const embeddedCaption = paragraph({
      index: 0,
      text: 'TABLE 1. Embedded',
      classification: 'tableCaption',
      tableNestingLevel: 1,
      containingTableIndex: 0,
      containingTableRowIndex: 0
    });
    const embedded = makeDocument(
      [embeddedCaption],
      [{ blockIndex: 0, kind: 'table', tableIndex: 0, tableCaptionRowIndex: 0 }]
    );
    expect(validateRule('TABLE_CAPTION_TABLE_GAP_001', embedded).status).toBe('not_applicable');
  });

  it('passes every safe table-caption gap paragraph to bottom-up deletion', async () => {
    const caption = paragraph({ index: 0, text: 'TABLE 1. Results', classification: 'tableCaption' });
    const makeBlank = (index: number) => paragraph({
      index,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const paragraphs = [caption, makeBlank(1), makeBlank(2)];
    const blocks: NonNullable<DocumentSnapshot['layout']['documentBlocks']> = [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, text: caption.text },
      { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1 },
      { blockIndex: 2, kind: 'paragraph', paragraphIndex: 2 },
      { blockIndex: 3, kind: 'table', tableIndex: 0 }
    ];
    const doc = snapshot(paragraphs);
    doc.layout.documentBlocks = blocks;
    doc.crossReferences = analyzeCrossReferences(paragraphs, blocks);
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(doc);
    const removeSpy = vi.spyOn(WordDocumentAdapter, 'removeBlankParagraphs').mockResolvedValue(undefined);

    try {
      await WordDocumentAdapter.applyTableCaptionTableGaps();
      expect(removeSpy).toHaveBeenCalledWith([1, 2]);
    } finally {
      snapshotSpy.mockRestore();
      removeSpy.mockRestore();
    }
  });

  it('detects non-single line spacing and locates offenders', () => {
    const finding = validateRule('LINE_SPACING_001', snapshot([
      paragraph({ index: 0 }),
      paragraph({ index: 1, lineSpacing: 18, lineSpacingMultiple: 1.5 })
    ]));

    expect(finding.status).toBe('failed');
    expect(finding.paragraphIndices).toEqual([1]);
    expect(finding.detected).toContain('1 paragraph(s) use non-single spacing');
    expect(finding.detected).toContain('Paragraph 2');
  });

  it('requires the OOXML auto+240 semantic representation for single spacing', () => {
    const auto240 = paragraph({ lineSpacing: 12, lineSpacingMultiple: 1, lineSpacingRule: 'auto', lineSpacingValue: 240 });
    const exact240 = paragraph({ lineSpacing: 12, lineSpacingMultiple: null, lineSpacingRule: 'exact', lineSpacingValue: 240 });
    const atLeast240 = paragraph({ lineSpacing: 12, lineSpacingMultiple: null, lineSpacingRule: 'atLeast', lineSpacingValue: 240 });
    const exact200 = paragraph({ lineSpacing: 10, lineSpacingMultiple: null, lineSpacingRule: 'exact', lineSpacingValue: 200 });

    expect(validateRule('LINE_SPACING_001', snapshot([auto240])).status).toBe('passed');
    const exactFinding = validateRule('LINE_SPACING_001', snapshot([exact240]));
    expect(exactFinding.status).toBe('failed');
    expect(exactFinding.autoFixable).toBe(true);
    expect(validateRule('LINE_SPACING_001', snapshot([atLeast240])).status).toBe('failed');
    expect(validateRule('LINE_SPACING_001', snapshot([exact200])).status).toBe('failed');
    const tableFinding = validateRule('LINE_SPACING_001', snapshot([
      paragraph({ index: 0, classification: 'tableCaption', lineSpacingRule: 'exact', lineSpacingMultiple: null, lineSpacingValue: 200 }),
      paragraph({ index: 1, classification: 'other', tableNestingLevel: 1, lineSpacingRule: 'exact', lineSpacingMultiple: null, lineSpacingValue: 200 })
    ]));
    expect(tableFinding.paragraphIndices).toEqual([0, 1]);
    expect(tableFinding.detected).toContain('table cell');
    expect(tableFinding.detected).toContain('200 twips');
  });

  it('verifies a single-spacing correction only after auto+240 is re-read', () => {
    const before = paragraph({ index: 7, lineSpacingRule: 'exact', lineSpacingMultiple: null, lineSpacingValue: 200 });
    const after = paragraph({ index: 7, lineSpacingRule: 'auto', lineSpacingMultiple: 1, lineSpacingValue: 240 });

    expect(isVerifiedSingleLineSpacing(before)).toBe(false);
    expect(verifySingleLineSpacing([before], [7])).toEqual([7]);
    expect(isVerifiedSingleLineSpacing(after)).toBe(true);
    expect(verifySingleLineSpacing([after], [7])).toEqual([]);
    expect(verifySingleLineSpacing([], [7])).toEqual([7]);
  });

  it('keeps confirmed formatting failures failed when another paragraph is unavailable', () => {
    const finding = validateRule('HEADING1_SIZE_001', snapshot([
      paragraph({ index: 0, classification: 'heading1', fontSize: 10 }),
      paragraph({ index: 1, classification: 'heading1', fontSize: null })
    ]));

    expect(finding.status).toBe('failed');
    expect(finding.paragraphIndices).toEqual([0]);
    expect(finding.detected).toContain('1 failed; 1 unavailable');
  });

  it('keeps confirmed line-spacing failures failed when another paragraph is unavailable', () => {
    const finding = validateRule('LINE_SPACING_001', snapshot([
      paragraph({ index: 0, lineSpacingMultiple: 1.5 }),
      paragraph({ index: 1, lineSpacingMultiple: null, lineSpacingRule: null })
    ]));

    expect(finding.status).toBe('failed');
    expect(finding.paragraphIndices).toEqual([0]);
    expect(finding.detected).toContain('1 failed; 1 unavailable');
  });

  it('keeps abstract prefix and period findings atomic', () => {
    const withPeriod = snapshot([paragraph({ classification: 'abstract', text: 'Abstract. Text' })]);
    const withoutPeriod = snapshot([paragraph({ classification: 'abstract', text: 'Abstract Text' })]);

    expect(validateRule('ABSTRACT_PREFIX_001', withPeriod).status).toBe('passed');
    expect(validateRule('ABSTRACT_PREFIX_PERIOD_001', withPeriod).status).toBe('passed');
    expect(validateRule('ABSTRACT_PREFIX_001', withoutPeriod).status).toBe('passed');
    expect(validateRule('ABSTRACT_PREFIX_PERIOD_001', withoutPeriod).status).toBe('failed');
  });

  it('offers a safe exact-range fix for an unbold Abstract. prefix', () => {
    const finding = validateRule('ABSTRACT_PREFIX_BOLD_001', snapshot([
      paragraph({
        index: 9,
        classification: 'abstract',
        text: 'Abstract. The body remains normal.',
        abstractPrefixBold: false,
        abstractBodyBold: false,
        ooxmlAligned: true
      })
    ]));
    expect(finding.status).toBe('failed');
    expect(finding.autoFixable).toBe(true);
    expect(finding.fixParagraphIndices).toEqual([9]);
  });

  it('passes an already bold Abstract. prefix and does not rewrite the body', () => {
    const finding = validateRule('ABSTRACT_PREFIX_BOLD_001', snapshot([
      paragraph({
        classification: 'abstract',
        text: 'Abstract. The body remains normal.',
        abstractPrefixBold: true,
        abstractBodyBold: false
      })
    ]));
    expect(finding.status).toBe('passed');
  });

  it('requires manual review for malformed or field-backed Abstract. prefixes', () => {
    const malformed = validateRule('ABSTRACT_PREFIX_BOLD_001', snapshot([
      paragraph({ classification: 'abstract', text: 'Abstract The body', abstractPrefixBold: false })
    ]));
    expect(malformed.status).toBe('manual_review');
    expect(malformed.autoFixable).toBe(false);

    const fieldBacked = validateRule('ABSTRACT_PREFIX_BOLD_001', snapshot([
      paragraph({
        classification: 'abstract',
        text: 'Abstract. The body',
        abstractPrefixBold: false,
        hasField: true,
        fieldTextSpans: [{ start: 0, end: 9 }]
      })
    ]));
    expect(fieldBacked.status).toBe('manual_review');
    expect(fieldBacked.autoFixable).toBe(false);
  });

  it('does not double-penalize heading minor words', () => {
    const lowercase = snapshot([paragraph({ classification: 'heading2', text: 'Analysis with Deep Learning' })]);
    const uppercaseMinor = snapshot([paragraph({ classification: 'heading2', text: 'Analysis With Deep Learning' })]);

    expect(validateRule('HEADING2_TITLE_CASE_001', lowercase).status).toBe('passed');
    expect(validateRule('HEADING2_MINOR_WORDS_001', lowercase).status).toBe('passed');
    expect(validateRule('HEADING2_TITLE_CASE_001', uppercaseMinor).status).toBe('passed');
    expect(validateRule('HEADING2_MINOR_WORDS_001', uppercaseMinor).status).toBe('failed');
  });

  it('accepts mixed-case technical terms in Heading 2 title capitalization', () => {
    expect(validateRule('HEADING2_TITLE_CASE_001', snapshot([
      paragraph({ classification: 'heading2', text: 'Comparison of Results with and Without mRMR Feature Selection' })
    ])).status).toBe('passed');
  });

  it('preserves technical terms such as XGBoost and mRMR in Heading 2', () => {
    expect(validateRule('HEADING2_TITLE_CASE_001', snapshot([
      paragraph({ classification: 'heading2', text: 'Prediction Using XGBoost and mRMR' })
    ])).status).toBe('passed');
  });

  it('requires Using to be capitalized as a major Heading 2 word', () => {
    expect(validateRule('HEADING2_TITLE_CASE_001', snapshot([
      paragraph({ classification: 'heading2', text: 'Prediction using XGBoost and mRMR' })
    ])).status).toBe('failed');
  });

  it('accepts eHealth and digit-bearing technical tokens in Heading 2', () => {
    expect(validateRule('HEADING2_TITLE_CASE_001', snapshot([
      paragraph({ classification: 'heading2', text: 'Applications of eHealth in 5G Networks' })
    ])).status).toBe('passed');
  });

  it('accepts scikit-learn as a technical token in Heading 2', () => {
    expect(validateRule('HEADING2_TITLE_CASE_001', snapshot([
      paragraph({ classification: 'heading2', text: 'Prediction Using scikit-learn Models' })
    ])).status).toBe('passed');
  });

  it('still rejects ordinary lowercase major words in Heading 2', () => {
    expect(validateRule('HEADING2_TITLE_CASE_001', snapshot([
      paragraph({ classification: 'heading2', text: 'comparison of Results with mRMR' })
    ])).status).toBe('failed');
  });

  it('does not treat arbitrary mixed-case typos as technical tokens', () => {
    expect(validateRule('HEADING2_TITLE_CASE_001', snapshot([
      paragraph({ classification: 'heading2', text: 'cOmparison of Results with mRMR' })
    ])).status).toBe('failed');
    expect(validateRule('HEADING2_TITLE_CASE_001', snapshot([
      paragraph({ classification: 'heading2', text: 'Analysis of reSults Using XGBoost' })
    ])).status).toBe('failed');
  });

  it('preserves complete technical tokens during title-case correction', () => {
    expect(toSafeTitleCase('Prediction using scikit-learn and XGBoost'))
      .toBe('Prediction Using scikit-learn and XGBoost');
  });

  it('accepts the LIM Heading 2 capitalization with Using', () => {
    const heading = 'Comparison of Results Using Different Classifiers';
    const doc = snapshot([paragraph({ classification: 'heading2', text: heading })]);
    expect(validateRule('HEADING2_TITLE_CASE_001', doc).status).toBe('passed');
    expect(validateRule('HEADING2_MINOR_WORDS_001', doc).status).toBe('passed');
  });

  it('treats Word All Caps formatting as visual capitalization', () => {
    const title = paragraph({
      classification: 'title',
      text: 'My Research Paper',
      allCaps: true
    });
    const heading = paragraph({
      classification: 'heading1',
      text: 'References',
      allCaps: true
    });

    expect(validateRule('TITLE_NOT_ALL_CAPS_001', snapshot([title])).status).toBe('failed');
    expect(validateRule('HEADING1_CAPS_001', snapshot([heading])).status).toBe('passed');
  });

  it('detects all-bold abstract body text without failing the prefix rule', () => {
    const finding = validateRule('ABSTRACT_BODY_NORMAL_001', snapshot([
      paragraph({ classification: 'abstract', abstractPrefixBold: true, abstractBodyBold: true })
    ]));

    expect(finding.status).toBe('failed');
    expect(finding.paragraphIndices).toEqual([0]);
  });

  it('fails mixed bold formatting in the remaining abstract text', () => {
    const finding = validateRule('ABSTRACT_BODY_NORMAL_001', snapshot([
      paragraph({ classification: 'abstract', abstractBodyBold: null, abstractBodyMixed: true })
    ]));

    expect(finding.status).toBe('failed');
    expect(finding.paragraphIndices).toEqual([0]);
  });

  it('marks Abstract body formatting unavailable only when it cannot be inspected', () => {
    const finding = validateRule('ABSTRACT_BODY_NORMAL_001', snapshot([
      paragraph({ classification: 'abstract', abstractBodyBold: null, abstractBodyMixed: null })
    ]));

    expect(finding.status).toBe('not_evaluated');
  });

  it('detects high-confidence AIP template boilerplate', () => {
    const finding = validateRule('TEMPLATE_BOILERPLATE_REMOVED_001', snapshot([
      paragraph({ text: 'How to Use this Template' })
    ]));

    expect(finding.status).toBe('failed');
    expect(finding.paragraphIndices).toEqual([0]);
  });

  it('classifies email text before the Author Affiliation style and excludes table cells from body', () => {
    const raw = (value: Partial<ParagraphSnapshot>) => {
      const { classification: _classification, ...withoutClassification } = paragraph(value);
      return withoutClassification;
    };
    const classified = WordDocumentAdapter.classify([
      raw({ index: 0, text: 'Paper title', style: 'Paper Title' }),
      raw({ index: 1, text: 'ari@example.edu', style: 'Author Affiliation' }),
      raw({ index: 2, text: 'University of Example', style: 'Author Affiliation' }),
      raw({ index: 3, text: 'AVE', tableNestingLevel: 1 }),
      raw({ index: 4, text: 'Abstract. Text', style: 'Abstract' })
    ]);

    expect(classified[1].classification).toBe('authorEmail');
    expect(classified[2].classification).toBe('authorAffiliation');
    expect(classified[3].classification).toBe('other');
  });

  it('classifies Faculty affiliations and Roman-numeral table captions semantically', () => {
    const classified = WordDocumentAdapter.classify([
      paragraph({ index: 0, text: 'Paper title' }),
      paragraph({ index: 1, text: 'John Smith' }),
      paragraph({ index: 2, text: '2Faculty of Vocational Studies, Universitas Brawijaya' }),
      paragraph({ index: 3, text: 'Abstract. Test', style: 'Abstract' }),
      paragraph({ index: 4, text: 'TABLE I. Comprehensive Comparison of Methods', style: 'BodyText' })
    ]);

    expect(classified[2].classification).toBe('authorAffiliation');
    expect(classified[4].classification).toBe('tableCaption');
    const cross = analyzeCrossReferences(classified, []);
    expect(cross.tableDefinitions).toEqual([
      expect.objectContaining({ number: 1, paragraphIndex: 4 })
    ]);
  });

  it('keeps text-box paragraphs out of the main-flow OOXML index sequence', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:t>Before</w:t></w:r></w:p>
      <w:p><w:r><w:drawing><w:txbxContent>
        <w:p><w:r><w:t>Nested A</w:t></w:r></w:p>
        <w:p><w:r><w:t>Nested B</w:t></w:r></w:p>
      </w:txbxContent></w:drawing></w:r></w:p>
      <w:p><w:r><w:t>After</w:t></w:r></w:p>
    </w:body></w:document>`;
    const stylesXml = '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';

    const paragraphs = parseParagraphOoxml(documentXml, stylesXml);
    expect(paragraphs.map(paragraph => paragraph.text)).toEqual(['Before', '', 'After']);
    expect(paragraphs.map(paragraph => paragraph.index)).toEqual([0, 1, 2]);

    const blocks = parseDocumentBlocks(documentXml);
    expect(blocks.filter(block => block.kind === 'paragraph').map(block => [block.paragraphIndex, block.text])).toEqual([
      [0, 'Before'],
      [1, ''],
      [2, 'After']
    ]);
  });

  it('protects dense tabbed drawing-based pseudo-table rows from prose formatting', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:t>TABLE 1. Results</w:t></w:r></w:p>
      <w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:t>Metric&#9;GNN&#9;SVM</w:t></w:r></w:p>
      <w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:t>Precision&#9;0.83&#9;0.85</w:t></w:r></w:p>
      <w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:t>Recall&#9;0.87&#9;0.89</w:t></w:r></w:p>
      <w:p><w:r><w:t>Normal prose resumes here.</w:t></w:r></w:p>
      <w:p><w:r><w:t>More ordinary discussion follows.</w:t></w:r></w:p>
      <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:t>GNN</w:t></w:r></w:p>
      <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:t>SVM</w:t></w:r></w:p>
    </w:body></w:document>`;
    const stylesXml = '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';
    const parsed = parseParagraphOoxml(documentXml, stylesXml);

    expect(parsed.slice(1, 4).every(paragraph => paragraph.layoutSensitive === true)).toBe(true);
    expect(parsed[1].layoutSignals?.hasCustomTabs).toBe(true);
    expect(parsed[4].layoutSensitive).not.toBe(true);
    expect(parsed[5].layoutSensitive).not.toBe(true);
    expect(parsed[6].layoutSensitive).not.toBe(true);
    expect(parsed[6].layoutSignals?.hasNumbering).toBe(true);
  });

  it('excludes protected pseudo-table rows from body and whitespace findings', () => {
    const protectedParagraph = paragraph({
      text: 'Metric  GNN  SVM',
      fontName: 'Calibri',
      layoutSensitive: true
    });
    expect(validateRule('BODY_FONT_001', snapshot([protectedParagraph])).status).toBe('not_applicable');
    expect(validateRule('TEXT_MULTIPLE_SPACES_001', snapshot([protectedParagraph])).status).toBe('passed');
  });

  it('uses one shared Arabic/Roman caption-number parser for table sequences', () => {
    expect(parseCaptionNumber('TABLE I. Methods', 'table')).toBe(1);
    expect(parseCaptionNumber('TABLE 2. Results', 'table')).toBe(2);
    const finding = validateRule('TABLE_NUMBER_SEQUENCE_001', snapshot([
      paragraph({ index: 0, text: 'TABLE I. Methods', classification: 'tableCaption', tablePrefixAllCaps: true }),
      paragraph({ index: 1, text: 'TABLE 2. Results', classification: 'tableCaption', tablePrefixAllCaps: true })
    ]));
    expect(finding.status).toBe('passed');
  });

  it('accepts only canonical Roman numerals for deterministic caption fixes', () => {
    expect(romanNumeralToNumber('IV')).toBe(4);
    expect(romanNumeralToNumber('IX')).toBe(9);
    expect(romanNumeralToNumber('XIV')).toBe(14);
    expect(romanNumeralToNumber('IIII')).toBeNull();
    expect(romanNumeralToNumber('IC')).toBeNull();
    expect(romanNumeralToNumber('VX')).toBeNull();
    expect(romanNumeralToNumber('IIV')).toBeNull();
    expect(parseCaptionNumber('TABLE IC. Invalid', 'table')).toBeNull();
  });

  it('fails closed for body validation when paragraph mapping counts disagree', () => {
    const doc = snapshot([paragraph({ fontName: 'Calibri' })]);
    doc.layout.paragraphMappingAligned = false;
    const finding = validateRule('BODY_FONT_001', doc);
    expect(finding.status).toBe('not_evaluated');
    expect(finding.reviewKind).toBe('technical');
  });

  it('does not claim BODY compliance when an individual paragraph mapping is unreliable', () => {
    const finding = validateRule('BODY_SIZE_001', snapshot([
      paragraph({ index: 0, fontSize: 10 }),
      paragraph({ index: 1, fontSize: 10, ooxmlAligned: false })
    ]));
    expect(finding.status).toBe('not_evaluated');
    expect(finding.paragraphIndices).toEqual([1]);
    expect(finding.detected).toContain('mapping unavailable');
  });

  it('does not score line spacing or repeated spaces on an unreliable paragraph', () => {
    const unreliable = paragraph({
      index: 1,
      ooxmlAligned: false,
      lineSpacingMultiple: 2,
      text: 'Body  paragraph'
    });
    const lineSpacing = validateRule('LINE_SPACING_001', snapshot([unreliable]));
    const spaces = validateRule('TEXT_MULTIPLE_SPACES_001', snapshot([unreliable]));

    expect(lineSpacing.status).toBe('not_evaluated');
    expect(spaces.status).toBe('not_evaluated');
    expect(lineSpacing.paragraphIndices).toEqual([1]);
    expect(spaces.paragraphIndices).toEqual([1]);
  });

  it('ignores layout-sensitive mismatches when building line-spacing and space-rule populations', () => {
    const pseudoTableRow = paragraph({
      index: 1,
      layoutSensitive: true,
      ooxmlAligned: false,
      lineSpacingMultiple: 2,
      text: 'Metric  GNN  SVM'
    });
    const lineSpacing = validateRule('LINE_SPACING_001', snapshot([pseudoTableRow]));
    const spaces = validateRule('TEXT_MULTIPLE_SPACES_001', snapshot([pseudoTableRow]));

    expect(lineSpacing.status).toBe('not_applicable');
    expect(spaces.status).toBe('passed');
    expect(lineSpacing.paragraphIndices).toEqual([]);
    expect(spaces.paragraphIndices).toEqual([]);
  });

  it('reports partially skipped line-spacing corrections', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const applySpacing = vi.spyOn(WordDocumentAdapter, 'applySingleLineSpacing').mockResolvedValue(undefined);
    const safety = snapshot([
      paragraph({ index: 0 }),
      paragraph({ index: 1, layoutSensitive: true })
    ]);
    const reread = snapshot([
      paragraph({ index: 0, lineSpacingRule: 'auto', lineSpacingMultiple: 1, lineSpacingValue: 240 }),
      paragraph({ index: 1, layoutSensitive: true })
    ]);
    vi.spyOn(WordDocumentAdapter, 'snapshot')
      .mockResolvedValueOnce(safety)
      .mockResolvedValueOnce(reread);
    (globalThis as { Word?: unknown }).Word = { run: vi.fn() };

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'LINE_SPACING_001',
          target: 'document',
          paragraphIndices: [0, 1],
          property: 'paragraph.lineSpacing',
          rawExpected: 12
        })
      ]);
      expect(applySpacing).toHaveBeenCalledWith([0]);
      expect(result.attempts[0].executionSucceeded).toBe(true);
      expect(result.attempts[0].message).toContain('1 paragraph(s) skipped');
    } finally {
      vi.restoreAllMocks();
      (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('resolves OOXML line spacing from direct and inherited style settings', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:before="240" w:after="240"/></w:pPr></w:p>
      <w:p><w:pPr><w:pStyle w:val="Wide"/></w:pPr></w:p>
      <w:p><w:pPr><w:spacing w:line="200" w:lineRule="exact"/></w:pPr></w:p>
      <w:p><w:pPr><w:pStyle w:val="SpacingOnly"/></w:pPr></w:p>
    </w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:spacing w:before="240" w:after="240"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Wide"><w:basedOn w:val="Base"/></w:style>
      <w:style w:type="paragraph" w:styleId="SpacingOnly"><w:basedOn w:val="Base"/><w:pPr><w:spacing w:before="240" w:after="240"/></w:pPr></w:style>
    </w:styles>`;
    const parsed = parseParagraphOoxml(documentXml, stylesXml);

    expect(parsed[0].lineSpacingMultiple).toBe(1);
    expect(parsed[1].lineSpacingMultiple).toBe(1.15);
    expect(parsed[2].lineSpacingMultiple).toBeNull();
    expect(parsed[2].lineSpacingRule).toBe('exact');
    expect(parsed[3].lineSpacingMultiple).toBe(1.15);
  });

  it('centralizes effective formatting values and their sources', () => {
    const styles = new Map([
      ['Child', { basedOn: 'Base', spaceBefore: 12 }],
      ['Base', { fontSize: 12, alignment: 'center', lineSpacing: { multiple: 1.15, rule: 'auto' } }]
    ]);
    const resolved = resolveEffectiveFormatting(
      { leftIndent: 18 },
      'Child',
      styles,
      { fontSize: 10, spaceAfter: 6, lineSpacing: { multiple: 1, rule: 'default' } }
    );

    expect(resolved.fontSize).toBe(12);
    expect(resolved.alignment).toBe('center');
    expect(resolved.leftIndent).toBe(18);
    expect(resolved.spaceBefore).toBe(12);
    expect(resolved.spaceAfter).toBe(6);
    expect(resolved.lineSpacingMultiple).toBe(1.15);
    expect(resolved.sources.leftIndent).toBe('direct');
    expect(resolved.sources.spaceBefore).toBe('style');
    expect(resolved.sources.fontSize).toBe('basedOn');
    expect(resolved.sources.spaceAfter).toBe('docDefaults');
  });

  it('resolves page-break-before through direct formatting, styles, basedOn, and explicit false', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Direct</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Break"/></w:pPr><w:r><w:t>Style</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="ChildBreak"/></w:pPr><w:r><w:t>Based on</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Break"/><w:pageBreakBefore w:val="0"/></w:pPr><w:r><w:t>Override</w:t></w:r></w:p>
    </w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Break"><w:pPr><w:pageBreakBefore/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="ChildBreak"><w:basedOn w:val="Break"/></w:style>
    </w:styles>`;

    const parsed = parseParagraphOoxml(documentXml, stylesXml);
    expect(parsed.map(paragraph => paragraph.hasPageBreakBefore)).toEqual([
      true,
      true,
      true,
      false
    ]);
  });

  it('uses docDefaults paragraph line spacing as the final OOXML fallback', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
    </w:styles>`;
    const parsed = parseParagraphOoxml(documentXml, stylesXml);

    expect(parsed[0].lineSpacingMultiple).toBe(1.15);
    expect(parsed[0].lineSpacingSource).toBe('docDefaults');
  });

  it('resolves rPrDefault, the default paragraph style, and direct run overrides', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:t>Normal text</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>Override</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>Styled</w:t></w:r></w:p>
    </w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults>
        <w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr></w:rPrDefault>
        <w:pPrDefault><w:pPr/></w:pPrDefault>
      </w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="Emphasis"><w:rPr><w:i/></w:rPr></w:style>
    </w:styles>`;
    const parsed = parseParagraphOoxml(documentXml, stylesXml);

    expect(parsed[0].fontName).toBe('Times New Roman');
    expect(parsed[0].fontSize).toBe(12);
    expect(parsed[0].formattingSources?.fontName).toBe('docDefaults');
    expect(parsed[1].fontName).toBe('Arial');
    expect(parsed[1].fontSize).toBe(12);
    expect(parsed[1].effectiveRuns?.[0].sources.fontName).toBe('direct');
    expect(parsed[2].effectiveRuns?.[0].italic).toBe(true);
    expect(parsed[2].effectiveRuns?.[0].sources.italic).toBe('style');
  });

  it('resolves superscript run metadata from OOXML', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r><w:r><w:t> Faculty</w:t></w:r></w:p></w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults></w:styles>`;
    const parsed = parseParagraphOoxml(documentXml, stylesXml);
    expect(parsed[0].effectiveRuns?.[0].verticalAlign).toBe('superscript');
  });

  it('resolves raised w:position run metadata without converting it to superscript', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:position w:val="10"/></w:rPr><w:t>1</w:t></w:r><w:r><w:t> Author</w:t></w:r></w:p></w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults></w:styles>`;
    const parsed = parseParagraphOoxml(documentXml, stylesXml);
    expect(parsed[0].effectiveRuns?.[0].baselinePositionPt).toBe(5);
    expect(parsed[0].effectiveRuns?.[0].verticalAlign).toBeNull();
  });

  it('resolves theme-based run fonts', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr><w:t>Theme text</w:t></w:r></w:p>
    </w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults></w:styles>`;
    const themeXml = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;
    const parsed = parseParagraphOoxml(documentXml, stylesXml, themeXml);

    expect(parsed[0].fontName).toBe('Aptos');
    expect(parsed[0].effectiveRuns?.[0].fontName).toBe('Aptos');
  });

  it('gives a theme font precedence over an explicit font in the same rFonts', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:asciiTheme="minorHAnsi"/></w:rPr><w:t>Theme wins</w:t></w:r></w:p>
    </w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults></w:styles>`;
    const themeXml = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;
    const parsed = parseParagraphOoxml(documentXml, stylesXml, themeXml);

    expect(parsed[0].effectiveRuns?.[0].fontName).toBe('Aptos');
  });

  it('fails a formatting rule when one text run directly overrides the paragraph style', () => {
    const finding = validateRule('BODY_SIZE_001', snapshot([
      paragraph({
        classification: 'body',
        // A mixed live Word range reports null; OOXML run values then provide
        // the authoritative fallback for the direct override.
        fontSize: null,
        effectiveRuns: [{
          text: 'Body with an override',
          fontName: 'Arial',
          fontSize: 12,
          bold: false,
          italic: false,
          allCaps: false,
          sources: { fontSize: 'direct' }
        }]
      })
    ]));

    expect(finding.status).toBe('failed');
    expect(finding.paragraphIndices).toEqual([0]);
  });

  it('uses the live uniform font after a correction instead of stale OOXML runs', () => {
    const finding = validateRule('TITLE_FONT_001', snapshot([
      paragraph({
        classification: 'title',
        fontName: 'Times New Roman',
        effectiveRuns: [{
          text: 'Corrected title',
          fontName: 'Calibri',
          fontSize: 18,
          bold: true,
          italic: false,
          allCaps: false,
          sources: { fontName: 'direct' }
        }]
      })
    ]));

    expect(finding.status).toBe('passed');
  });

  it('falls back to effective runs when live font formatting is mixed', () => {
    const finding = validateRule('BODY_FONT_001', snapshot([
      paragraph({
        classification: 'body',
        fontName: null,
        effectiveRuns: [
          {
            text: 'Body one',
            fontName: 'Times New Roman',
            fontSize: 10,
            bold: false,
            italic: false,
            allCaps: false,
            sources: { fontName: 'style' }
          },
          {
            text: ' body two',
            fontName: 'Calibri',
            fontSize: 10,
            bold: false,
            italic: false,
            allCaps: false,
            sources: { fontName: 'direct' }
          }
        ]
      })
    ]));

    expect(finding.status).toBe('failed');
  });

  it('uses the live bad font even when stale OOXML says the expected font', () => {
    const finding = validateRule('TITLE_FONT_001', snapshot([
      paragraph({
        classification: 'title',
        fontName: 'Calibri',
        effectiveRuns: [{
          text: 'Uncorrected title',
          fontName: 'Times New Roman',
          fontSize: 18,
          bold: true,
          italic: false,
          allCaps: false,
          sources: { fontName: 'style' }
        }]
      })
    ]));

    expect(finding.status).toBe('failed');
  });

  it('does not treat a References table cell as the References heading', () => {
    const classified = WordDocumentAdapter.classify([
      paragraph({ index: 0, text: 'Perfect Research Paper' }),
      paragraph({ index: 1, text: 'Abstract. Test', style: 'Abstract' }),
      paragraph({ index: 2, text: 'References', tableNestingLevel: 1 }),
      paragraph({ index: 3, text: 'CONCLUSION', style: 'Heading 1' }),
      paragraph({ index: 4, text: 'Conclusion body.', style: 'Paragraph' }),
      paragraph({ index: 5, text: 'REFERENCES', style: 'Heading 1', tableNestingLevel: 0 }),
      paragraph({ index: 6, text: '1. First reference', style: 'Reference' })
    ]);

    expect(classified[2].classification).toBe('other');
    expect(classified[5].classification).toBe('referencesHeading');
    expect(classified[6].classification).toBe('reference');
    expect(classified[4].classification).toBe('body');
  });

  it('recognizes an embedded table caption and associates it with its containing table', () => {
    const raw = (value: Partial<ParagraphSnapshot>) => {
      const { classification: _classification, ...withoutClassification } = paragraph(value);
      return withoutClassification;
    };
    const classified = WordDocumentAdapter.classify([
      raw({ index: 0, text: 'Paper title' }),
      raw({ index: 1, text: 'Abstract. Test', style: 'Abstract' }),
      raw({ index: 2, text: 'TABLE 1. Results', tableNestingLevel: 1 })
    ]);
    expect(classified[2].classification).toBe('tableCaption');

    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>TABLE 1. Results</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Data</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const parsed = parseParagraphOoxml(documentXml, '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>');
    expect(parsed[0].containingTableIndex).toBe(0);
    expect(parsed[0].containingTableRowIndex).toBe(0);
    const blocks = parseDocumentBlocks(documentXml);
    expect(blocks[0].tableCaptionRowIndex).toBe(0);
    expect(blocks[0].tableHeaderStartRowIndex).toBe(1);
    expect(blocks[0].tableHeaderEndRowIndex).toBe(1);

    const caption = paragraph({
      index: 0,
      text: 'TABLE 1. Results',
      classification: 'tableCaption',
      tableNestingLevel: 1,
      containingTableIndex: 0,
      containingTableRowIndex: 0
    });
    const cross = analyzeCrossReferences([caption], blocks);
    expect(cross.tablePlacement).toEqual([
      expect.objectContaining({ objectBlockIndex: 0, passed: true, confidence: 'high' })
    ]);
  });

  it('resolves style-inherited reference numbering and separates label punctuation from the tab suffix', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Reference"/></w:pPr><w:r><w:t>Author, Example paper (2025).</w:t></w:r></w:p>
    </w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Paragraph"><w:name w:val="Paragraph"/></w:style>
      <w:style w:type="paragraph" w:styleId="Reference"><w:name w:val="Reference"/><w:basedOn w:val="Paragraph"/><w:next w:val="Reference"/>
        <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>
      </w:style>
    </w:styles>`;
    const numberingXml = `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="tab"/><w:lvlJc w:val="left"/>
        <w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr>
      </w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num>
    </w:numbering>`;
    const parsed = parseParagraphOoxml(documentXml, stylesXml, '', numberingXml)[0];

    expect(parsed.referenceNumbering).toEqual(expect.objectContaining({
      kind: 'automatic',
      source: 'style',
      level: 1,
      format: 'decimal',
      suffix: '.',
      separator: 'tab',
      alignment: 'left',
      numberPositionPt: 18,
      textPositionPt: 36
    }));
  });

  it('classifies high-confidence unnumbered bibliography entries after REFERENCES', () => {
    const raw = (value: Partial<ParagraphSnapshot>) => {
      const { classification: _classification, ...withoutClassification } = paragraph(value);
      return withoutClassification;
    };
    const classified = WordDocumentAdapter.classify([
      raw({ index: 0, text: 'Paper title' }),
      raw({ index: 1, text: 'Abstract. Summary', style: 'Abstract' }),
      raw({ index: 2, text: 'REFERENCES', style: 'Heading 1', styleBuiltIn: 'Heading 1' }),
      raw({ index: 3, text: 'A. Author, Example research paper, Journal Name 12, 25–30 (2025).', style: 'Normal' }),
      raw({ index: 4, text: 'B. Author, Another research paper, Proceedings Title (2026).', style: 'Normal' })
    ]);

    expect(classified.slice(2).map(item => item.classification)).toEqual([
      'referencesHeading', 'reference', 'reference'
    ]);
  });

  it('lets explicit TABLE/FIGURE text override conflicting caption styles', () => {
    const classified = WordDocumentAdapter.classify([
      paragraph({ index: 0, text: 'Paper title' }),
      paragraph({ index: 1, text: 'TABLE 1. Algorithm', style: 'Figure Caption' }),
      paragraph({ index: 2, text: 'FIGURE 1. Architecture', style: 'Table Caption' }),
      paragraph({ index: 3, text: 'Architecture overview', style: 'Figure Caption' }),
      paragraph({ index: 4, text: 'TABLE I. Algorithm', style: 'Figure Caption' }),
      paragraph({ index: 5, text: 'TABLE. Results', style: 'Figure Caption' }),
      paragraph({ index: 6, text: 'TABLE Results', style: 'Figure Caption' })
    ]);

    expect(classified[1].classification).toBe('tableCaption');
    expect(classified[2].classification).toBe('figureCaption');
    expect(classified[3].classification).toBe('figureCaption');
    expect(classified[4].classification).toBe('tableCaption');
    expect(classified[5].classification).toBe('tableCaption');
    expect(classified[6].classification).toBe('tableCaption');
    expect(hasExplicitCaptionPrefix('TABLE 1. Algorithm', 'figure')).toBe(false);
    expect(hasExplicitCaptionPrefix('TABLE 1. Algorithm', 'table')).toBe(true);
  });

  it('extracts only the leading figure caption label', () => {
    expect(extractFigureCaptionLabel('Fig. 1. Performance of ResNet50')).toBe('Fig. 1.');
    expect(extractFigureCaptionLabel('FIG. 2: Results')).toBe('FIG. 2:');
    expect(extractFigureCaptionLabel('figure 3) Architecture')).toBe('figure 3)');
    expect(extractFigureCaptionLabel('Architecture overview')).toBeNull();
    expect(extractCaptionPrefix('TABLE I. Methods', 'table')).toBe('TABLE I.');
  });

  it('refuses to rewrite a malformed TABLE caption passed to the figure fixer', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const paragraph = {
      text: 'TABLE I. Algorithm',
      load: vi.fn(),
      getRange: vi.fn()
    };
    const collection = {
      items: [paragraph],
      load: vi.fn()
    };
    const context = {
      document: { body: { paragraphs: collection } },
      sync: vi.fn().mockResolvedValue(undefined)
    };

    (globalThis as { Word?: unknown }).Word = {
      run: async (callback: (value: typeof context) => Promise<void>) => callback(context),
      RangeLocation: { content: 'Content' },
      InsertLocation: { start: 'Start', replace: 'Replace' }
    };

    try {
      await WordDocumentAdapter.normalizeFigureCaptionPrefixes([0], [0]);
      expect(paragraph.text).toBe('TABLE I. Algorithm');
      expect(paragraph.getRange).not.toHaveBeenCalled();
    } finally {
      (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('changes only a figure caption prefix and preserves the caption body', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const bodyText = 'Performance of ResNet50';
    const paragraph = {
      text: `Fig. 1. ${bodyText}`,
      load: vi.fn(),
      getRange: vi.fn()
    } as {
      text: string;
      load: ReturnType<typeof vi.fn>;
      getRange: ReturnType<typeof vi.fn>;
    };
    const prefixRange = {
      font: { bold: false },
      getOoxml: vi.fn(() => ({ value: '<w:r><w:t>Fig. 1.</w:t></w:r>' })),
      insertText: vi.fn((replacement: string) => {
        paragraph.text = `${replacement}${paragraph.text.slice('Fig. 1.'.length)}`;
        return { font: { bold: false } };
      })
    };
    const contentRange = {
      search: vi.fn(() => ({ items: [prefixRange], load: vi.fn() }))
    };
    paragraph.getRange.mockReturnValue(contentRange);

    const collection = {
      items: [paragraph],
      load: vi.fn()
    };
    const context = {
      document: { body: { paragraphs: collection } },
      sync: vi.fn().mockResolvedValue(undefined)
    };

    (globalThis as { Word?: unknown }).Word = {
      run: async (callback: (value: typeof context) => Promise<void>) => callback(context),
      RangeLocation: { content: 'Content' },
      InsertLocation: { start: 'Start', replace: 'Replace' }
    };

    try {
      await WordDocumentAdapter.normalizeFigureCaptionPrefixes([0], [0]);
      expect(prefixRange.insertText).toHaveBeenCalledWith('FIGURE 1.', 'Replace');
      expect(paragraph.text).toBe(`FIGURE 1. ${bodyText}`);
      expect(paragraph.text.endsWith(bodyText)).toBe(true);
    } finally {
      (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('resolves All Caps from OOXML styles and individual runs', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>References</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:b/><w:caps/></w:rPr><w:t>Figure 1.</w:t></w:r><w:r><w:t> Research Model</w:t></w:r></w:p>
      <w:p><w:r><w:t>MY </w:t></w:r><w:r><w:rPr><w:caps/></w:rPr><w:t>title</w:t></w:r></w:p>
    </w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Heading1"><w:rPr><w:caps/></w:rPr></w:style>
    </w:styles>`;
    const parsed = parseParagraphOoxml(documentXml, stylesXml);

    expect(parsed[0].allCaps).toBe(true);
    expect(parsed[1].allCaps).toBe(false);
    expect(parsed[1].figurePrefixAllCaps).toBe(true);
    expect(parsed[2].allCaps).toBe(true);
  });

  it('parses the supplied AIP template as a real DOCX fixture', async () => {
    const fixture = await readFile(new URL('../reference/AIP Template(7).docx', import.meta.url));
    const zip = await JSZip.loadAsync(fixture);
    const documentXml = await zip.file('word/document.xml')?.async('string');
    const stylesXml = await zip.file('word/styles.xml')?.async('string') || '';
    const numberingXml = await zip.file('word/numbering.xml')?.async('string') || '';

    expect(documentXml).toBeTruthy();
    const parsed = parseParagraphOoxml(documentXml || '', stylesXml, '', numberingXml);
    const nonSingle = parsed.filter(p => p.lineSpacingMultiple !== 1 || p.lineSpacingRule === 'exact' || p.lineSpacingRule === 'atLeast');

    expect(parsed.length).toBeGreaterThan(100);
    expect(nonSingle).toHaveLength(0);
    expect(parsed.some(p => p.formattingSources?.lineSpacingMultiple === 'default')).toBe(true);
    expect(parsed.filter(p => p.fontName === 'Times New Roman').length).toBeGreaterThan(100);
    expect(parsed.some(p => p.formattingSources?.fontName === 'docDefaults')).toBe(true);

    const references = parsed.filter(p => p.referenceStyle?.name === 'Reference');
    expect(references.length).toBeGreaterThan(0);
    expect(references.every(p => p.referenceStyle?.leftIndent === 0)).toBe(true);
    expect(references.every(p => p.referenceStyle?.nextParagraphStyle === null)).toBe(true);
    expect(references.every(p => Math.abs((p.referenceStyle?.hangingIndent || 0) - 21.3) <= 0.75)).toBe(true);
    const automaticReferences = references.filter(p => p.referenceNumbering?.kind === 'automatic');
    expect(automaticReferences.length).toBeGreaterThan(0);
    expect(automaticReferences.every(p => p.referenceNumbering?.format === 'decimal')).toBe(true);
    expect(automaticReferences.every(p => p.referenceNumbering?.suffix === '.')).toBe(true);
    expect(automaticReferences.every(p => p.referenceNumbering?.numberPositionPt === 18)).toBe(true);
    expect(automaticReferences.every(p => p.referenceNumbering?.textPositionPt === 36)).toBe(true);
  });

  it('recognizes the AIP template embedded table-caption arrangement', async () => {
    const fixture = await readFile(new URL('../reference/AIP Template(7).docx', import.meta.url));
    const zip = await JSZip.loadAsync(fixture);
    const documentXml = await zip.file('word/document.xml')?.async('string');
    const stylesXml = await zip.file('word/styles.xml')?.async('string') || '';
    expect(documentXml).toBeTruthy();
    const tableBlocks = parseDocumentBlocks(documentXml || '', stylesXml).filter(block => block.kind === 'table');
    const aipTable = tableBlocks.find(block =>
      block.tableCaptionRowIndex === 0 && block.tableHeaderStartRowIndex === 1
    );
    expect(aipTable).toBeTruthy();
    expect(aipTable?.tableHeaderEndRowIndex).toBe(1);
    expect(aipTable?.tableTopBorder).toBe(true);
    expect(aipTable?.tableBottomBorder).toBe(true);
    expect(aipTable?.tableHeaderRule).toBe(true);
    expect(aipTable?.tableVerticalRules).toBe(false);
    expect(aipTable?.tableExtraHorizontalRules).toBe(false);
  });

  it('rejects an embedded table caption that is not in the first logical row', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>Data before caption</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>TABLE 1. Late caption</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Data</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const blocks = parseDocumentBlocks(documentXml);
    expect(blocks[0].tableCaptionRowIndex).toBeNull();
    expect(blocks[0].tableHeaderStartRowIndex).toBe(0);
    const cross = analyzeCrossReferences([paragraph({
      index: 1,
      text: 'TABLE 1. Late caption',
      classification: 'tableCaption',
      tableNestingLevel: 1,
      containingTableIndex: 0,
      containingTableRowIndex: 1
    })], blocks);
    expect(cross.tablePlacement[0]).toEqual(expect.objectContaining({
      passed: false,
      confidence: 'high',
      objectBlockIndex: 0
    }));
  });

  it('detects an extra horizontal boundary between multi-row header rows', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:tcBorders><w:bottom w:val="single"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>Group</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Column</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const table = parseDocumentBlocks(documentXml)[0];
    expect(table.tableHeaderStartRowIndex).toBe(0);
    expect(table.tableHeaderEndRowIndex).toBe(1);
    expect(table.tableHeaderRule).toBe(false);
    expect(table.tableExtraHorizontalRules).toBe(true);
  });

  it('keeps caption-presence checks applicable when an object has no caption', () => {
    const figure = validateRule('FIGURE_CAPTION_PRESENCE_001', {
      ...snapshot([]),
      layout: {
        ...snapshot([]).layout,
        documentBlocks: [{
          blockIndex: 0,
          kind: 'paragraph',
          paragraphIndex: 0,
          drawingCount: 1,
          hasInlineDrawing: true,
          hasAnchoredDrawing: false,
          text: ''
        }]
      }
    });
    const table = validateRule('TABLE_CAPTION_PRESENCE_001', {
      ...snapshot([]),
      layout: {
        ...snapshot([]).layout,
        documentBlocks: [{
          blockIndex: 0,
          kind: 'table',
          tableIndex: 0,
          drawingCount: 0,
          text: 'Header Value'
        }]
      }
    });

    expect(figure.status).toBe('failed');
    expect(figure.detected).toContain('1 figure/drawing candidate');
    expect(table.status).toBe('failed');
    expect(table.detected).toContain('1 table candidate');
  });

  it('ignores equation-layout tables and pairs a caption with the immediate manuscript table', () => {
    const paragraphs = [
      paragraph({ index: 0, text: 'TABLE 2. Results', classification: 'tableCaption' }),
      paragraph({ index: 1, text: 'TABLE 3. Results', classification: 'tableCaption' })
    ];
    const blocks = [
      { blockIndex: 0, kind: 'paragraph' as const, paragraphIndex: 0, text: paragraphs[0].text },
      { blockIndex: 1, kind: 'table' as const, tableIndex: 0, isEquationLayout: false, drawingCount: 0, text: 'Actual table' },
      { blockIndex: 2, kind: 'table' as const, tableIndex: 1, isEquationLayout: true, drawingCount: 0, text: '(1)' },
      { blockIndex: 3, kind: 'table' as const, tableIndex: 2, isEquationLayout: true, drawingCount: 0, text: '(2)' },
      { blockIndex: 4, kind: 'paragraph' as const, paragraphIndex: 1, text: paragraphs[1].text }
    ];
    const doc = snapshot(paragraphs);
    doc.layout.documentBlocks = blocks;
    doc.crossReferences = analyzeCrossReferences(paragraphs, blocks);

    expect(doc.crossReferences.tablePlacement[0]).toEqual(expect.objectContaining({
      objectBlockIndex: 1,
      confidence: 'high',
      passed: true
    }));
    expect(validateRule('TABLE_CAPTION_PRESENCE_001', doc).status).toBe('passed');
  });

  it('infers a multi-row header from vertical merges when tblHeader metadata is absent', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Measure</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Commands</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>ON</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>OFF</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Accuracy</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>0</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const table = parseDocumentBlocks(documentXml)[0];
    expect(table.tableHeaderStartRowIndex).toBe(0);
    expect(table.tableHeaderEndRowIndex).toBe(1);
    expect(table.tableHeaderRowCount).toBe(2);
  });

  it('distinguishes a complete required border from a partial forbidden border', () => {
    const partialHeaderXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:tcBorders><w:bottom w:val="single"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>3</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>4</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const partialHeader = parseDocumentBlocks(partialHeaderXml)[0];
    expect(partialHeader.tableHeaderRule).toBe(false);

    const partialBodyXml = partialHeaderXml.replace(
      '<w:tr><w:tc><w:p><w:r><w:t>3</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>4</w:t></w:r></w:p></w:tc></w:tr>',
      '<w:tr><w:tc><w:tcPr><w:tcBorders><w:bottom w:val="single"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>3</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>4</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>5</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>6</w:t></w:r></w:p></w:tc></w:tr>'
    );
    const partialBody = parseDocumentBlocks(partialBodyXml)[0];
    expect(partialBody.tableExtraHorizontalRules).toBe(true);
  });

  it('treats table-level insideH as extra across multi-row headers', () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tblPr><w:tblBorders><w:insideH w:val="single"/></w:tblBorders></w:tblPr>
      <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Group</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Column</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    expect(parseDocumentBlocks(xml)[0].tableExtraHorizontalRules).toBe(true);
  });

  it('resolves table top and bottom borders inherited from a table style', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="table" w:styleId="TableGrid"><w:tblPr><w:tblBorders>
        <w:top w:val="single"/><w:bottom w:val="single"/>
      </w:tblBorders></w:tblPr></w:style>
    </w:styles>`;
    const table = parseDocumentBlocks(documentXml, stylesXml)[0];

    expect(table.tableStyleId).toBe('TableGrid');
    expect(table.effectiveTableBorders).toEqual(expect.objectContaining({ top: true, bottom: true }));
    expect(table.tableTopBorder).toBe(true);
    expect(table.tableBottomBorder).toBe(true);
  });

  it('lets a direct table nil border override an inherited table-style border', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblBorders><w:top w:val="nil"/></w:tblBorders></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="table" w:styleId="TableGrid"><w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr></w:style>
    </w:styles>`;
    const table = parseDocumentBlocks(documentXml, stylesXml)[0];

    expect(table.effectiveTableBorders?.top).toBe(false);
    expect(table.tableTopBorder).toBe(false);
  });

  it('resolves table borders through a basedOn table-style chain', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tblPr><w:tblStyle w:val="CustomGrid"/></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="table" w:styleId="TableGrid"><w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr></w:style>
      <w:style w:type="table" w:styleId="CustomGrid"><w:basedOn w:val="TableGrid"/><w:tblPr><w:tblBorders><w:bottom w:val="single"/></w:tblBorders></w:tblPr></w:style>
    </w:styles>`;
    const table = parseDocumentBlocks(documentXml, stylesXml)[0];

    expect(table.effectiveTableBorders).toEqual(expect.objectContaining({ top: true, bottom: true }));
    expect(table.effectiveTableBorders?.sources?.top).toBe('style:TableGrid');
    expect(table.effectiveTableBorders?.sources?.bottom).toBe('style:CustomGrid');
  });

  it('merges direct table-style borders with wholeTable conditional borders', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tblPr><w:tblStyle w:val="HybridGrid"/></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="table" w:styleId="HybridGrid">
        <w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr>
        <w:tblStylePr w:type="wholeTable"><w:tblPr><w:tblBorders><w:insideV w:val="single"/></w:tblBorders></w:tblPr></w:tblStylePr>
      </w:style>
    </w:styles>`;
    const table = parseDocumentBlocks(documentXml, stylesXml)[0];

    expect(table.effectiveTableBorders?.top).toBe(true);
    expect(table.effectiveTableBorders?.insideV).toBe(true);
  });

  it('does not apply a firstRow conditional border to the whole table', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tblPr><w:tblStyle w:val="ConditionalOnly"/></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="table" w:styleId="ConditionalOnly">
        <w:tblStylePr w:type="firstRow"><w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr></w:tblStylePr>
      </w:style>
    </w:styles>`;
    const table = parseDocumentBlocks(documentXml, stylesXml)[0];

    expect(table.effectiveTableBorders?.top).toBeNull();
    expect(table.tableTopBorder).toBe(false);
  });

  it('treats an inherited insideV border as a forbidden vertical rule', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="table" w:styleId="TableGrid"><w:tblPr><w:tblBorders><w:insideV w:val="single"/></w:tblBorders></w:tblPr></w:style>
    </w:styles>`;
    const table = parseDocumentBlocks(documentXml, stylesXml)[0];

    expect(table.effectiveTableBorders?.insideV).toBe(true);
    expect(table.tableVerticalRules).toBe(true);
  });

  it('uses the composite top edge when cell borders override a table-level nil value', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tblPr><w:tblBorders><w:top w:val="nil"/></w:tblBorders></w:tblPr>
      <w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:tcBorders><w:top w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc></w:tr>
      <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const table = parseDocumentBlocks(documentXml)[0];
    expect(table.effectiveTableBorders?.top).toBe(false);
    expect(table.tableTopBorder).toBe(true);

    const caption = paragraph({ index: 0, text: 'TABLE 1. Results', classification: 'tableCaption' });
    const doc = snapshot([caption]);
    doc.layout.documentBlocks = [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, text: caption.text },
      { ...table, blockIndex: 1 }
    ];
    doc.crossReferences = {
      referenceDefinitions: [], referenceCitations: [], referenceCitationStyle: 'unknown',
      figureDefinitions: [], figureCitations: [], tableDefinitions: [{ kind: 'table', number: 1, paragraphIndex: 0 }],
      tableCitations: [], referenceCoverageMissing: [], invalidReferenceCitations: [], referenceSequenceViolations: [],
      figureCoverageMissing: [], tableCoverageMissing: [], figurePlacement: [],
      tablePlacement: [{ kind: 'table', number: 1, captionParagraphIndex: 0, objectBlockIndex: 1, passed: true, confidence: 'high' }],
      figurePriorCitationFailures: [], figurePriorCitationUnknown: []
    };

    expect(validateRule('TABLE_TOP_RULE_001', doc).status).toBe('passed');
  });

  it('uses the composite vertical result when cell borders override table-level nil values', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tblPr><w:tblBorders><w:insideV w:val="nil"/><w:left w:val="nil"/><w:right w:val="nil"/></w:tblBorders></w:tblPr>
      <w:tr><w:tc><w:tcPr><w:tcBorders><w:left w:val="single"/><w:right w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
      <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const table = parseDocumentBlocks(documentXml)[0];
    expect(table.effectiveTableBorders?.insideV).toBe(false);
    expect(table.tableVerticalRules).toBe(true);

    const caption = paragraph({ index: 0, text: 'TABLE 1. Results', classification: 'tableCaption' });
    const doc = snapshot([caption]);
    doc.layout.documentBlocks = [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, text: caption.text },
      { ...table, blockIndex: 1 }
    ];
    doc.crossReferences = {
      referenceDefinitions: [], referenceCitations: [], referenceCitationStyle: 'unknown',
      figureDefinitions: [], figureCitations: [], tableDefinitions: [{ kind: 'table', number: 1, paragraphIndex: 0 }],
      tableCitations: [], referenceCoverageMissing: [], invalidReferenceCitations: [], referenceSequenceViolations: [],
      figureCoverageMissing: [], tableCoverageMissing: [], figurePlacement: [],
      tablePlacement: [{ kind: 'table', number: 1, captionParagraphIndex: 0, objectBlockIndex: 1, passed: true, confidence: 'high' }],
      figurePriorCitationFailures: [], figurePriorCitationUnknown: []
    };

    expect(validateRule('TABLE_NO_VERTICAL_RULES_001', doc).status).toBe('failed');
  });

  it('uses insideH as the required top boundary for an embedded caption', () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tblPr><w:tblBorders><w:insideH w:val="single"/><w:bottom w:val="single"/></w:tblBorders></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t>TABLE 1. Caption</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    const table = parseDocumentBlocks(xml)[0];
    expect(table.tableCaptionRowIndex).toBe(0);
    expect(table.tableTopBorder).toBe(true);
    expect(table.tableHeaderRule).toBe(true);
    expect(table.tableExtraHorizontalRules).toBe(false);
  });

  it('shows the multi-part figure review for medium-confidence layout tables', () => {
    const doc = snapshot([
      paragraph({ index: 0, text: 'FIGURE 1. Results', classification: 'figureCaption' })
    ]);
    doc.layout.documentBlocks = [{
      blockIndex: 7,
      kind: 'table',
      tableIndex: 0,
      drawingCount: 2,
      text: ''
    }];
    doc.crossReferences = {
      referenceDefinitions: [], referenceCitations: [], referenceCitationStyle: 'unknown',
      figureDefinitions: [{ kind: 'figure', number: 1, paragraphIndex: 0 }],
      figureCitations: [], tableDefinitions: [], tableCitations: [],
      referenceCoverageMissing: [], invalidReferenceCitations: [], referenceSequenceViolations: [],
      figureCoverageMissing: [], tableCoverageMissing: [], tablePlacement: [],
      figurePlacement: [{ kind: 'figure', number: 1, captionParagraphIndex: 0, objectBlockIndex: 7, passed: true, confidence: 'medium' }],
      figurePriorCitationFailures: [], figurePriorCitationUnknown: []
    };
    expect(validateRule('FIGURE_SUBPART_LABELS_001', doc).status).toBe('manual_review');
  });

  it('does not advertise unsupported automatic correction properties', () => {
    expect(strategy.rules.find(rule => rule.id === 'PAGE_WIDTH_001')?.autoFixable).toBe(true);
    expect(strategy.rules.find(rule => rule.id === 'TABLE_TEXT_CITATION_001')?.autoFixable).toBe(false);
    expect(strategy.rules.find(rule => rule.id === 'CONCLUSION_BODY_FONT_001')?.autoFixable).toBe(true);
    const dedicatedRuleIds = new Set([
      'HEADER_FORBIDDEN_001', 'FOOTER_FORBIDDEN_001', 'LINE_SPACING_001',
      'MARGIN_TOP_001', 'MARGIN_RIGHT_001', 'MARGIN_BOTTOM_001', 'MARGIN_LEFT_001', 'PAGE_NUMBER_FORBIDDEN_001',
      'TEXT_MULTIPLE_SPACES_001', 'TITLE_CASE_001', 'FIGURE_CAPTION_PREFIX_001',
      'FIGURE_CAPTION_LEADING_WHITESPACE_001', 'TABLE_CAPTION_LEADING_WHITESPACE_001',
      'REFERENCES_NUMBERING_001', 'AUTHOR_AFFILIATION_SUPERSCRIPT_001',
      'AUTHOR_EMAIL_SUPERSCRIPT_001', 'AUTHOR_AFFILIATION_ITALIC_001',
      'AUTHOR_EMAIL_ITALIC_001', 'EQUATION_TABS_001', 'TABLE_TOP_RULE_001',
      'TABLE_BOTTOM_RULE_001', 'TABLE_HEADER_RULE_001', 'TABLE_NO_VERTICAL_RULES_001',
      'TABLE_NO_EXTRA_HORIZONTAL_RULES_001', 'EXTRA_BLANK_PARAGRAPH_001',
      'PAGE_WIDTH_001', 'PAGE_HEIGHT_001', 'FIGURE_CAPTION_BODY_GAP_001', 'FIGURE_CITATION_SPELLING_001',
      'FIGURE_LINE_SPACE_ABOVE_001', 'TABLE_CAPTION_TABLE_GAP_001', 'CONCLUSION_BODY_FONT_001',
      'HEADING1_STYLE_001', 'FIRST_HEADING1_SPACE_BEFORE_001', 'AUTHOR_NAME_SPACING_001',
      'AUTHOR_AFFILIATION_SPACING_001', 'AUTHOR_EMAIL_SPACING_001', 'ABSTRACT_SPACING_001',
      'AFFILIATION_EMAIL_BLANK_001', 'FRONTMATTER_EXTRA_BLANK_001', 'ABSTRACT_PREFIX_BOLD_001',
      'TABLE_PREFIX_BOLD_001', 'REFERENCE_FONT_001', 'REFERENCE_SIZE_001',
      'REFERENCE_ALIGNMENT_001', 'REFERENCE_STYLE_001', 'REFERENCE_LIST_LEVEL_001',
      'REFERENCE_NUMBER_POSITION_001', 'REFERENCE_TEXT_POSITION_001',
      'REFERENCES_HANGING_INDENT_001', 'REFERENCES_LEFT_INDENT_001'
    ]);
    const genericProperties = new Set([
      'font.name', 'font.size', 'font.bold', 'font.italic', 'paragraph.alignment',
      'paragraph.firstLineIndent', 'paragraph.leftIndent', 'paragraph.rightIndent',
      'paragraph.spaceBefore', 'paragraph.spaceAfter', 'text.allCaps'
    ]);
    const unmapped = strategy.rules.filter(rule => rule.autoFixable
      && !dedicatedRuleIds.has(rule.id)
      && !genericProperties.has(rule.property));
    expect(unmapped).toEqual([]);
  });

  it('applies Heading 1 style through a dedicated fresh-snapshot correction', async () => {
    const finding = validateRule('HEADING1_STYLE_001', snapshot([
      paragraph({
        index: 5,
        text: 'INTRODUCTION',
        classification: 'body',
        style: 'Normal',
        sectionFamilies: ['introduction'],
        sectionConfidence: 'high',
        heading1Confidence: 'high',
        ooxmlAligned: true
      })
    ]));
    const applyStyle = vi.spyOn(WordDocumentAdapter, 'applyBuiltInParagraphStyle').mockResolvedValue(undefined);
    const refreshed = paragraph({
      index: 5,
      text: 'INTRODUCTION',
      classification: 'body',
      style: 'Localized Heading 1',
      styleBuiltIn: 'Heading1',
      sectionFamilies: ['introduction'],
      sectionConfidence: 'high',
      heading1Confidence: 'high',
      ooxmlAligned: true
    });
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot')
      .mockResolvedValueOnce(snapshot([paragraph({
        index: 5,
        text: 'INTRODUCTION',
        classification: 'body',
        style: 'Normal',
        sectionFamilies: ['introduction'],
        sectionConfidence: 'high',
        heading1Confidence: 'high',
        ooxmlAligned: true
      })]))
      .mockResolvedValueOnce(snapshot([refreshed]));
    try {
      const result = await CorrectionService.apply([finding]);
      expect(applyStyle).toHaveBeenCalledWith([5], 'Heading1');
      expect(result.attempts[0].executionSucceeded).toBe(true);
      expect(result.executedFixes).toBe(1);
      expect(snapshotSpy).toHaveBeenCalledTimes(2);
    } finally {
      applyStyle.mockRestore();
      snapshotSpy.mockRestore();
    }
  });

  it('refuses a stale Heading 1 style correction when the validated text changes', async () => {
    const finding = validateRule('HEADING1_STYLE_001', snapshot([
      paragraph({
        index: 6,
        text: 'INTRODUCTION',
        classification: 'body',
        style: 'Normal',
        sectionFamilies: ['introduction'],
        sectionConfidence: 'high',
        heading1Confidence: 'high',
        ooxmlAligned: true
      })
    ]));
    const applyStyle = vi.spyOn(WordDocumentAdapter, 'applyBuiltInParagraphStyle').mockResolvedValue(undefined);
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(snapshot([
      paragraph({
        index: 6,
        text: 'Changed paragraph',
        classification: 'body',
        style: 'Normal',
        sectionFamilies: [],
        sectionConfidence: 'low',
        ooxmlAligned: true
      })
    ]));
    try {
      const result = await CorrectionService.apply([finding]);
      expect(applyStyle).not.toHaveBeenCalled();
      expect(result.executedFixes).toBe(0);
      expect(result.executionFailures).toBe(1);
      expect(result.attempts[0].message).toContain('text or layout mapping changed');
    } finally {
      applyStyle.mockRestore();
      snapshotSpy.mockRestore();
    }
  });

  it('uses one-to-one object associations for caption presence', () => {
    const base = snapshot([
      paragraph({ index: 0, text: 'FIGURE 1. First', classification: 'figureCaption' }),
      paragraph({ index: 1, text: 'FIGURE 1. Duplicate', classification: 'figureCaption' })
    ]);
    base.layout.documentBlocks = [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 2, drawingCount: 1, hasInlineDrawing: true },
      { blockIndex: 1, kind: 'paragraph', paragraphIndex: 3, drawingCount: 1, hasInlineDrawing: true }
    ];
    base.crossReferences = {
      referenceDefinitions: [], referenceCitations: [], referenceCitationStyle: 'unknown',
      figureDefinitions: [
        { kind: 'figure', number: 1, paragraphIndex: 0 },
        { kind: 'figure', number: 1, paragraphIndex: 1 }
      ],
      figureCitations: [], tableDefinitions: [], tableCitations: [],
      referenceCoverageMissing: [], invalidReferenceCitations: [], referenceSequenceViolations: [],
      figureCoverageMissing: [], tableCoverageMissing: [], tablePlacement: [],
      figurePlacement: [{ kind: 'figure', number: 1, captionParagraphIndex: 0, objectBlockIndex: 0, passed: true, confidence: 'high' }],
      figurePriorCitationFailures: [], figurePriorCitationUnknown: []
    };
    const finding = validateRule('FIGURE_CAPTION_PRESENCE_001', base);
    expect(finding.status).toBe('failed');
    expect(finding.detected).toContain('1 figure/drawing candidate');
    expect(finding.assessmentRole).toBe('consistency');
  });

  it('deduplicates table fixes without shifting later table metadata', () => {
    const paragraphs = [
      paragraph({ index: 0, text: 'TABLE 1. First', classification: 'tableCaption' }),
      paragraph({ index: 1, text: 'TABLE 1. Duplicate', classification: 'tableCaption' }),
      paragraph({ index: 2, text: 'TABLE 2. Second', classification: 'tableCaption' })
    ];
    const doc = snapshot(paragraphs);
    doc.layout.documentBlocks = [
      { blockIndex: 10, kind: 'table', tableIndex: 0, tableHeaderRule: false, tableHeaderRowCount: 2, tableHeaderEndRowIndex: 1 },
      { blockIndex: 11, kind: 'table', tableIndex: 1, tableHeaderRule: false, tableHeaderRowCount: 4, tableHeaderEndRowIndex: 3 }
    ];
    doc.crossReferences = {
      referenceDefinitions: [], referenceCitations: [], referenceCitationStyle: 'unknown',
      figureDefinitions: [], figureCitations: [], tableDefinitions: [
        { kind: 'table', number: 1, paragraphIndex: 0 },
        { kind: 'table', number: 1, paragraphIndex: 1 },
        { kind: 'table', number: 2, paragraphIndex: 2 }
      ],
      tableCitations: [], referenceCoverageMissing: [], invalidReferenceCitations: [], referenceSequenceViolations: [],
      figureCoverageMissing: [], tableCoverageMissing: [], figurePlacement: [],
      tablePlacement: [
        { kind: 'table', number: 1, captionParagraphIndex: 0, objectBlockIndex: 10, passed: false, confidence: 'high' },
        { kind: 'table', number: 1, captionParagraphIndex: 1, objectBlockIndex: 10, passed: false, confidence: 'high' },
        { kind: 'table', number: 2, captionParagraphIndex: 2, objectBlockIndex: 11, passed: false, confidence: 'high' }
      ],
      figurePriorCitationFailures: [], figurePriorCitationUnknown: []
    };
    const finding = validateRule('TABLE_HEADER_RULE_001', doc);
    expect(finding.paragraphIndices).toEqual([0, 2]);
    expect(finding.fixTableIndices).toEqual([0, 1]);
    expect(finding.fixTableHeaderRowCounts).toEqual([2, 4]);
  });

  it('extracts Figure/Table prefixes regardless of capitalization', () => {
    expect(extractCaptionPrefix('Figure 1. Research Model', 'figure')).toBe('Figure 1.');
    expect(extractCaptionPrefix('table 2. Results', 'table')).toBe('table 2.');
  });

  it('keeps figure sequence separate from prefix capitalization', () => {
    const paragraphs = [
      paragraph({ index: 0, classification: 'figureCaption', text: 'Figure 1. First' }),
      paragraph({ index: 1, classification: 'figureCaption', text: 'Figure 2. Second' })
    ];
    expect(validateRule('FIGURE_NUMBER_SEQUENCE_001', snapshot(paragraphs)).status).toBe('passed');
    expect(validateRule('FIGURE_CAPTION_PREFIX_001', snapshot([paragraphs[0]])).status).toBe('failed');
  });

  it('downgrades a field-backed figure caption label to manual review', () => {
    const fieldBacked = paragraph({
      index: 0,
      classification: 'figureCaption',
      text: 'Fig. 1. Results',
      fieldTextSpans: [{ start: 5, end: 6 }]
    });
    const finding = validateRule('FIGURE_CAPTION_PREFIX_001', snapshot([fieldBacked]));

    expect(finding.status).toBe('manual_review');
    expect(finding.autoFixable).toBe(false);
    expect(finding.reviewKind).toBe('manual');
    expect(finding.message).toContain('Word field');
  });

  it('keeps ordinary figure prefix fixes available beside a field-backed label', () => {
    const safe = paragraph({
      index: 0,
      classification: 'figureCaption',
      text: 'Fig. 1. Results'
    });
    const fieldBacked = paragraph({
      index: 1,
      classification: 'figureCaption',
      text: 'Fig. 2. Architecture',
      fieldTextSpans: [{ start: 5, end: 6 }]
    });
    const finding = validateRule('FIGURE_CAPTION_PREFIX_001', snapshot([safe, fieldBacked]));

    expect(finding.status).toBe('failed');
    expect(finding.autoFixable).toBe(true);
    expect(finding.paragraphIndices).toEqual([0, 1]);
    expect(finding.detected).toContain('field-backed label');
  });

  it('downgrades malformed existing figure labels instead of offering a duplicate prefix', () => {
    for (const text of ['FIGURE I. Results', 'FIGURE. Results', 'Fig. I. Architecture']) {
      const finding = validateRule('FIGURE_CAPTION_PREFIX_001', snapshot([
        paragraph({ index: 0, classification: 'figureCaption', text })
      ]));

      expect(finding.status).toBe('manual_review');
      expect(finding.autoFixable).toBe(false);
      expect(finding.reviewKind).toBe('manual');
      expect(finding.detected).toContain('malformed');
    }
  });

  it('keeps a genuinely unlabeled figure caption fixable', () => {
    const finding = validateRule('FIGURE_CAPTION_PREFIX_001', snapshot([
      paragraph({ index: 0, classification: 'figureCaption', text: 'Results of the proposed architecture' })
    ]));

    expect(finding.status).toBe('failed');
    expect(finding.autoFixable).toBe(true);
  });

  it('does not prepend a second label to a malformed figure caption', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const paragraph = {
      text: 'FIGURE I. Results',
      load: vi.fn(),
      getRange: vi.fn()
    };
    const collection = {
      items: [paragraph],
      load: vi.fn()
    };
    const context = {
      document: { body: { paragraphs: collection } },
      sync: vi.fn().mockResolvedValue(undefined)
    };

    (globalThis as { Word?: unknown }).Word = {
      run: async (callback: (value: typeof context) => Promise<void>) => callback(context),
      RangeLocation: { content: 'Content' },
      InsertLocation: { start: 'Start', replace: 'Replace' }
    };

    try {
      await WordDocumentAdapter.normalizeFigureCaptionPrefixes([0], [0]);
      expect(paragraph.text).toBe('FIGURE I. Results');
      expect(paragraph.getRange).not.toHaveBeenCalled();
    } finally {
      (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('accepts Figure and Table prefixes rendered with Word All Caps', () => {
    const figure = paragraph({
      classification: 'figureCaption',
      text: 'Figure 1. Research Model',
      figurePrefixBold: true,
      figurePrefixAllCaps: true
    });
    const table = paragraph({
      classification: 'tableCaption',
      text: 'Table 1. Results',
      tablePrefixBold: true,
      tablePrefixAllCaps: true
    });

    expect(validateRule('FIGURE_CAPTION_PREFIX_001', snapshot([figure])).status).toBe('passed');
    expect(validateRule('TABLE_CAPTION_PREFIX_001', snapshot([table])).status).toBe('passed');
    expect(validateRule('TABLE_PREFIX_CAPS_001', snapshot([table])).status).toBe('passed');
  });

  it('offers a safe exact-range fix for an unbold TABLE prefix', () => {
    const finding = validateRule('TABLE_PREFIX_BOLD_001', snapshot([
      paragraph({
        index: 0,
        classification: 'tableCaption',
        text: 'TABLE 1. Results',
        tablePrefixBold: false
      })
    ]));

    expect(finding.status).toBe('failed');
    expect(finding.autoFixable).toBe(true);
    expect(finding.fixParagraphIndices).toEqual([0]);
  });

  it('does not offer a TABLE prefix bold fix when the label is field-backed', () => {
    const finding = validateRule('TABLE_PREFIX_BOLD_001', snapshot([
      paragraph({
        index: 0,
        classification: 'tableCaption',
        text: 'TABLE 1. Results',
        tablePrefixBold: false,
        fieldTextSpans: [{ start: 6, end: 7 }]
      })
    ]));

    expect(finding.status).toBe('manual_review');
    expect(finding.autoFixable).toBe(false);
  });

  it('validates References as a Heading 1 equivalent', () => {
    const references = paragraph({
      classification: 'referencesHeading',
      text: 'References',
      spaceBefore: 12,
      spaceAfter: 12,
      bold: true,
      fontSize: 12,
      alignment: 'centered',
      allCaps: true
    });
    expect(validateRule('REFERENCES_HEADING_CAPS_001', snapshot([references])).status).toBe('passed');
    expect(validateRule('REFERENCES_HEADING_SPACE_BEFORE_001', snapshot([references])).status).toBe('passed');
    expect(validateRule('REFERENCES_HEADING_SPACE_AFTER_001', snapshot([references])).status).toBe('passed');
  });

  it('does not turn unavailable effective formatting into a false failure', () => {
    const heading = paragraph({ classification: 'heading1', fontSize: null });
    const finding = validateRule('HEADING1_SIZE_001', snapshot([heading]));

    expect(finding.status).toBe('not_evaluated');
    expect(finding.status).not.toBe('failed');
  });

  it('expands numeric reference citation lists and ranges', () => {
    expect(extractReferenceCitations('Prior work [1], [1, 3] and [4-6].')).toEqual([[1], [1, 3], [4, 5, 6]]);
    expect(extractReferenceCitations('A range [1–3, 6] is supported.')).toEqual([[1, 2, 3, 6]]);
  });

  it('detects reference coverage, invalid citations, and first-use sequence', () => {
    const analyzed = analyzeCrossReferences([
      paragraph({ index: 0, text: 'As shown in [1].', classification: 'body' }),
      paragraph({ index: 1, text: 'A later citation [3].', classification: 'body' }),
      paragraph({ index: 2, text: 'REFERENCES', classification: 'referencesHeading' }),
      paragraph({ index: 3, text: '1. First reference', classification: 'reference' }),
      paragraph({ index: 4, text: '2. Second reference', classification: 'reference' }),
      paragraph({ index: 5, text: '3. Third reference', classification: 'reference' })
    ], [], 2);

    expect(analyzed.referenceCoverageMissing).toEqual([2]);
    expect(analyzed.referenceSequenceViolations.map(item => item.paragraphIndex)).toEqual([1]);
    expect(analyzed.invalidReferenceCitations).toHaveLength(0);

    const invalid = analyzeCrossReferences([
      paragraph({ index: 0, text: 'Unsupported [9].', classification: 'body' }),
      paragraph({ index: 1, text: 'REFERENCES', classification: 'referencesHeading' }),
      paragraph({ index: 2, text: '1. First reference', classification: 'reference' })
    ], [], 1);
    expect(invalid.invalidReferenceCitations.map(item => item.paragraphIndex)).toEqual([0]);
    expect(invalid.referenceSequenceViolations).toHaveLength(0);
  });

  it('resolves Word automatic-list reference numbers', () => {
    const analyzed = analyzeCrossReferences([
      paragraph({ index: 0, text: 'Cited in [1].', classification: 'body' }),
      paragraph({ index: 1, text: 'REFERENCES', classification: 'referencesHeading' }),
      paragraph({ index: 2, text: 'Brown and Smith, Example.', classification: 'reference', isListItem: true, listString: '1.' })
    ], [], 1);

    expect(analyzed.referenceDefinitions).toEqual([
      { kind: 'reference', number: 1, paragraphIndex: 2 }
    ]);
    expect(analyzed.referenceCoverageMissing).toEqual([]);
  });

  it('validates Reference style and automatic-list layers independently', () => {
    const goodReference = paragraph({
      index: 0,
      text: 'Author, Example.',
      style: 'Reference',
      classification: 'reference',
      firstLineIndent: -21.6,
      isListItem: true,
      listString: '1.',
      referenceStyle: {
        name: 'Reference',
        type: 'paragraph',
        basedOn: 'Paragraph',
        nextParagraphStyle: 'Reference',
        fontName: 'Times New Roman',
        fontSize: 10,
        alignment: 'justified',
        leftIndent: 0,
        hangingIndent: 21.6
      },
      referenceNumbering: {
        kind: 'automatic',
        level: 1,
        format: 'decimal',
        suffix: '.',
        alignment: 'left',
        startAt: 1,
        numberPositionPt: 18,
        textPositionPt: 36
      }
    });
    expect(validateRule('REFERENCES_NUMBERING_001', snapshot([goodReference])).status).toBe('passed');
    expect(validateRule('REFERENCES_HANGING_INDENT_001', snapshot([goodReference])).status).toBe('passed');
    expect(validateRule('REFERENCES_LEFT_INDENT_001', snapshot([goodReference])).status).toBe('passed');
    expect(validateRule('REFERENCE_FONT_001', snapshot([goodReference])).status).toBe('passed');
    expect(validateRule('REFERENCE_SIZE_001', snapshot([goodReference])).status).toBe('passed');
    expect(validateRule('REFERENCE_ALIGNMENT_001', snapshot([goodReference])).status).toBe('passed');
    expect(validateRule('REFERENCE_STYLE_001', snapshot([goodReference])).status).toBe('passed');
    expect(validateRule('REFERENCE_LIST_LEVEL_001', snapshot([goodReference])).status).toBe('passed');
    expect(validateRule('REFERENCE_NUMBER_POSITION_001', snapshot([goodReference])).status).toBe('passed');
    expect(validateRule('REFERENCE_TEXT_POSITION_001', snapshot([goodReference])).status).toBe('passed');

    const wrongParagraphStyle = paragraph({
      ...goodReference,
      style: 'Normal',
      styleBuiltIn: 'Normal',
      referenceStyle: undefined
    });
    expect(validateRule('REFERENCE_STYLE_001', snapshot([wrongParagraphStyle])).status).toBe('failed');
    expect(validateRule('REFERENCE_FONT_001', snapshot([wrongParagraphStyle])).status).toBe('failed');

    const directFormattingOverride = paragraph({
      ...goodReference,
      fontName: 'Calibri',
      fontSize: 12,
      alignment: 'left'
    });
    expect(validateRule('REFERENCE_FONT_001', snapshot([directFormattingOverride])).status).toBe('failed');
    expect(validateRule('REFERENCE_SIZE_001', snapshot([directFormattingOverride])).status).toBe('failed');
    expect(validateRule('REFERENCE_ALIGNMENT_001', snapshot([directFormattingOverride])).status).toBe('failed');

    const mixedRunFormatting = paragraph({
      ...goodReference,
      effectiveRuns: [
        { text: 'Author', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, sources: {} },
        { text: 'Article title', fontName: 'Calibri', fontSize: 12, bold: false, italic: false, allCaps: false, sources: {} }
      ]
    });
    expect(validateRule('REFERENCE_FONT_001', snapshot([mixedRunFormatting])).status).toBe('failed');
    expect(validateRule('REFERENCE_SIZE_001', snapshot([mixedRunFormatting])).status).toBe('failed');

    const manualReference = paragraph({
      ...goodReference,
      text: '1. Author, Example.',
      isListItem: false,
      listString: null,
      referenceNumbering: { ...goodReference.referenceNumbering!, kind: 'manual' }
    });
    expect(validateRule('REFERENCES_NUMBERING_001', snapshot([manualReference])).status).toBe('failed');

    const wrongNumberAlignment = paragraph({
      ...goodReference,
      referenceNumbering: { ...goodReference.referenceNumbering!, alignment: 'right' }
    });
    expect(validateRule('REFERENCES_NUMBERING_001', snapshot([wrongNumberAlignment])).status).toBe('failed');
  });

  it('accepts the AIP raw style indentation and omitted implicit next style', () => {
    const templateReference = paragraph({
      ...paragraph({
        style: 'Reference',
        classification: 'reference',
        firstLineIndent: -21.3,
        isListItem: true,
        listString: '1.'
      }),
      referenceStyle: {
        name: 'Reference',
        type: 'paragraph',
        basedOn: 'Paragraph',
        nextParagraphStyle: null,
        fontName: 'Times New Roman',
        fontSize: 10,
        alignment: 'justified',
        leftIndent: 21.3,
        hangingIndent: 21.3
      },
      referenceNumbering: {
        kind: 'automatic',
        level: 1,
        format: 'decimal',
        suffix: '.',
        alignment: 'left',
        startAt: 1,
        numberPositionPt: 18,
        textPositionPt: 36
      }
    });

    expect(validateRule('REFERENCES_LEFT_INDENT_001', snapshot([templateReference])).status).toBe('passed');
    expect(validateRule('REFERENCE_STYLE_001', snapshot([templateReference])).status).toBe('passed');

    const explicitlyNormalNext = paragraph({
      ...templateReference,
      referenceStyle: { ...templateReference.referenceStyle!, nextParagraphStyle: 'Normal' }
    });
    expect(validateRule('REFERENCE_STYLE_001', snapshot([explicitlyNormalNext])).status).toBe('failed');

    const directIndentOverride = paragraph({
      ...templateReference,
      formattingSources: { leftIndent: 'direct', firstLineIndent: 'direct' }
    });
    expect(validateRule('REFERENCES_LEFT_INDENT_001', snapshot([directIndentOverride])).status).toBe('failed');
  });

  it('normalizes raw AIP Reference style indentation to the user-facing left anchor', () => {
    const parsed = parseParagraphOoxml(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Reference"/></w:pPr><w:r><w:t>Reference text</w:t></w:r></w:p></w:body></w:document>',
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Reference"><w:name w:val="Reference"/><w:basedOn w:val="Paragraph"/><w:pPr><w:ind w:left="426" w:hanging="426"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="20"/></w:rPr></w:style></w:styles>'
    );
    expect(parsed[0].referenceStyle?.leftIndent).toBe(0);
    expect(parsed[0].referenceStyle?.hangingIndent).toBeCloseTo(21.3);
  });

  it('reports missing figure numbers and preserves manual-review status', () => {
    const captions = [
      ...Array.from({ length: 16 }, (_, index) => index + 1),
      19, 20, 21
    ].map((number, index) => paragraph({
      index,
      text: `FIGURE ${number}. Caption`,
      classification: 'figureCaption'
    }));
    const finding = validateRule('FIGURE_NUMBER_SEQUENCE_001', snapshot(captions));
    expect(finding.status).toBe('failed');
    expect(finding.detected).toContain('Missing figure numbers: 17, 18');
    expect(finding.detected).toContain('Paragraph 17: expected 17, detected 19');
    expect(finding.detected).toContain('Paragraph 18: expected 18, detected 20');
    expect(finding.detected).toContain('Paragraph 19: expected 19, detected 21');
    expect(finding.paragraphIndices).toEqual([16, 17, 18]);
    expect(finding.autoFixable).toBe(false);

    const document = snapshot([
      paragraph({ index: 0, text: 'Figure 1 is shown.', classification: 'body' }),
      paragraph({ index: 1, text: 'Figure 19 is shown later.', classification: 'body' }),
      ...captions.map((caption, index) => ({ ...caption, index: index + 2 }))
    ]);
    document.crossReferences = analyzeCrossReferences(document.paragraphs);
    const citationFinding = validateRule('FIGURE_CITATION_SEQUENCE_001', document);
    expect(citationFinding.status).toBe('failed');
    expect(citationFinding.detected).toContain('Missing figure numbers');
    expect(citationFinding.autoFixable).toBe(false);

    const outOfOrder = validateRule('FIGURE_NUMBER_SEQUENCE_001', snapshot([
      paragraph({ index: 0, text: 'FIGURE 1. Caption', classification: 'figureCaption' }),
      paragraph({ index: 1, text: 'FIGURE 3. Caption', classification: 'figureCaption' }),
      paragraph({ index: 2, text: 'FIGURE 2. Caption', classification: 'figureCaption' })
    ]));
    expect(outOfOrder.status).toBe('failed');
    expect(outOfOrder.detected).toBe('Figure captions appear out of numerical order');
  });

  it('does not count a caption as its own figure/table citation', () => {
    const analyzed = analyzeCrossReferences([
      paragraph({ index: 0, text: 'FIGURE 1. Results', classification: 'figureCaption' }),
      paragraph({ index: 1, text: 'FIGURE 2. More results', classification: 'figureCaption' }),
      paragraph({ index: 2, text: 'TABLE 1. Values', classification: 'tableCaption' }),
      paragraph({ index: 3, text: 'TABLE 2. More values', classification: 'tableCaption' }),
      paragraph({ index: 4, text: 'The result is shown in Figs. 1 and 2 and Tables 1 and 2.', classification: 'body' })
    ]);

    expect(analyzed.figureCoverageMissing).toEqual([]);
    expect(analyzed.tableCoverageMissing).toEqual([]);
    expect(analyzed.figureCitations[0].paragraphIndex).toBe(4);
    expect(analyzed.tableCitations[0].paragraphIndex).toBe(4);
  });

  it('does not count figure citations found only inside table cells', () => {
    const analyzed = analyzeCrossReferences([
      paragraph({ index: 0, text: 'See Fig. 1.', classification: 'other', tableNestingLevel: 1 }),
      paragraph({ index: 1, text: 'FIGURE 1. Results', classification: 'figureCaption' })
    ]);

    expect(analyzed.figureCitations).toEqual([]);
    expect(analyzed.figureCoverageMissing).toEqual([1]);
  });

  it('does not count table citations found only inside table cells', () => {
    const analyzed = analyzeCrossReferences([
      paragraph({ index: 0, text: 'See Table 1.', classification: 'other', tableNestingLevel: 1 }),
      paragraph({ index: 1, text: 'TABLE 1. Results', classification: 'tableCaption' })
    ]);

    expect(analyzed.tableCitations).toEqual([]);
    expect(analyzed.tableCoverageMissing).toEqual([1]);
  });

  it('parses plural object citations with commas, and, ranges, and subfigure suffixes', () => {
    const analyzed = analyzeCrossReferences([
      paragraph({ index: 0, text: 'Figures 1, 2, and 3; Figs. 4–6; Fig. 7A.', classification: 'body' }),
      ...[1, 2, 3, 4, 5, 6, 7].map(number => paragraph({
        index: number,
        text: `FIGURE ${number}. Caption`,
        classification: 'figureCaption'
      }))
    ]);

    expect(analyzed.figureCoverageMissing).toEqual([]);
  });

  it('parses plural parenthesized subfigure citations', () => {
    const analyzed = analyzeCrossReferences([
      paragraph({ index: 0, text: 'The comparison is shown in Figs. 1(a) and 2(b).', classification: 'body' }),
      paragraph({ index: 1, text: 'FIGURE 1. First result', classification: 'figureCaption' }),
      paragraph({ index: 2, text: 'FIGURE 2. Second result', classification: 'figureCaption' })
    ]);

    expect(analyzed.figureCoverageMissing).toEqual([]);
  });

  it('checks first figure-citation order while ignoring repeated mentions', () => {
    const captions = [1, 2, 3].map(number => paragraph({
      index: number + 1,
      text: `FIGURE ${number}. Caption`,
      classification: 'figureCaption'
    }));
    const outOfOrder = analyzeCrossReferences([
      paragraph({ index: 0, text: 'Figure 2 is shown first.', classification: 'body' }),
      paragraph({ index: 1, text: 'Figure 1 is discussed next.', classification: 'body' }),
      ...captions
    ]);
    expect(outOfOrder.figureCitationSequenceViolations?.map(item => item.paragraphIndex)).toEqual([0]);

    const inOrder = analyzeCrossReferences([
      paragraph({ index: 0, text: 'Figures 1 and 2 are compared.', classification: 'body' }),
      paragraph({ index: 1, text: 'Figure 1 is revisited, then Figure 3.', classification: 'body' }),
      ...captions
    ]);
    expect(inOrder.figureCitationSequenceViolations).toEqual([]);

    const doc = snapshot([
      paragraph({ index: 0, text: 'Figure 2 is shown first.', classification: 'body' }),
      ...captions
    ]);
    doc.crossReferences = analyzeCrossReferences(doc.paragraphs);
    expect(validateRule('FIGURE_CITATION_SEQUENCE_001', doc).status).toBe('failed');
    expect(validateRule('FIGURE_CITATION_SEQUENCE_001', doc).autoFixable).toBe(false);
  });

  it('requires full Figure/Figures wording and exposes prefix-only fixes', () => {
    const occurrences = extractFigureCitationOccurrences(
      'Figure 1 shows the result. Shown in Figure 2. e.g. Figure 3 illustrates. Figs. 4 and 5 show more.',
      0
    );
    expect(occurrences.map(item => ({
      original: item.originalText,
      expected: item.expectedPrefix,
      sentenceStart: item.isSentenceStart
    }))).toEqual([
      { original: 'Figure', expected: 'Figure', sentenceStart: true },
      { original: 'Figure', expected: 'Figure', sentenceStart: false },
      { original: 'Figure', expected: 'Figure', sentenceStart: false },
      { original: 'Figs.', expected: 'Figures', sentenceStart: true }
    ]);

    expect(occurrences[2].originalCitationText).toBe('Figure 3');
    expect(occurrences[2].correctedCitationText).toBe('Figure 3');

    const wrappedAbbreviationOccurrences = extractFigureCitationOccurrences(
      '(e.g. Figure 6 illustrates the result) “e.g. Figure 7 illustrates another result”',
      0
    );
    expect(wrappedAbbreviationOccurrences.map(item => item.expectedPrefix)).toEqual(['Figure', 'Figure']);

    const paragraphs = [
      paragraph({ index: 0, text: 'Figs. 1 and 2 are compared. See Fig. 3.', classification: 'body' }),
      paragraph({ index: 1, text: 'Figure 1. Results', classification: 'figureCaption' })
    ];
    const doc = snapshot(paragraphs);
    doc.crossReferences = analyzeCrossReferences(paragraphs);
    const finding = validateRule('FIGURE_CITATION_SPELLING_001', doc);
    expect(finding.status).toBe('failed');
    expect(finding.paragraphIndices).toEqual([0]);
    expect(finding.fixTextRanges).toEqual([{
      paragraphIndex: 0,
      searchText: 'Figs. 1 and 2',
      originalPrefix: 'Figs.',
      replacementPrefix: 'Figures',
      replacement: 'Figures 1 and 2',
      occurrenceIndex: 0
    }, {
      paragraphIndex: 0,
      searchText: 'Fig. 3',
      originalPrefix: 'Fig.',
      replacementPrefix: 'Figure',
      replacement: 'Figure 3',
      occurrenceIndex: 0
    }]);
    expect(finding.autoFixable).toBe(true);
  });

  it('keeps field-backed figure citation spelling as manual review', () => {
    const paragraphs = [
      paragraph({ index: 0, text: 'See Fig. 1 and Fig. 2.', classification: 'body', fieldTextSpans: [{ start: 4, end: 21 }] }),
      paragraph({ index: 1, text: 'FIGURE 1. First result', classification: 'figureCaption' }),
      paragraph({ index: 2, text: 'FIGURE 2. Second result', classification: 'figureCaption' })
    ];
    const doc = snapshot(paragraphs);
    doc.crossReferences = analyzeCrossReferences(paragraphs);
    const finding = validateRule('FIGURE_CITATION_SPELLING_001', doc);
    expect(finding.status).toBe('manual_review');
    expect(finding.autoFixable).toBe(false);
    expect(finding.reviewKind).toBe('manual');
    expect(finding.fixTextRanges).toBeUndefined();
  });

  it('keeps safe figure citation fixes when another paragraph is field-backed', () => {
    const paragraphs = [
      paragraph({ index: 0, text: 'See Fig. 1 and Fig. 2.', classification: 'body', fieldTextSpans: [{ start: 16, end: 19 }] }),
      paragraph({ index: 1, text: 'FIGURE 1. First result', classification: 'figureCaption' }),
      paragraph({ index: 2, text: 'FIGURE 2. Second result', classification: 'figureCaption' })
    ];
    const doc = snapshot(paragraphs);
    doc.crossReferences = analyzeCrossReferences(paragraphs);
    const finding = validateRule('FIGURE_CITATION_SPELLING_001', doc);
    expect(finding.status).toBe('failed');
    expect(finding.autoFixable).toBe(true);
    expect(finding.fixTextRanges).toEqual([{
      paragraphIndex: 0,
      searchText: 'Fig. 1',
      originalPrefix: 'Fig.',
      replacementPrefix: 'Figure',
      replacement: 'Figure 1',
      occurrenceIndex: 0
    }]);
  });

  it('does not let an unrelated field block a figure citation fix', () => {
    const paragraphs = [
      paragraph({ index: 0, text: 'See Fig. 2 on page 5.', classification: 'body', fieldTextSpans: [{ start: 16, end: 17 }] }),
      paragraph({ index: 1, text: 'FIGURE 2. Result', classification: 'figureCaption' })
    ];
    const doc = snapshot(paragraphs);
    doc.crossReferences = analyzeCrossReferences(paragraphs);
    const finding = validateRule('FIGURE_CITATION_SPELLING_001', doc);
    expect(finding.status).toBe('failed');
    expect(finding.autoFixable).toBe(true);
    expect(finding.fixTextRanges?.map(range => range.searchText)).toEqual(['Fig. 2']);
  });

  it('maps a REF field result to the exact figure citation span', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
      <w:r><w:t>See Figure </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText> REF _Ref123 </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>2</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:t>.</w:t></w:r>
    </w:p></w:body></w:document>`;
    const parsed = parseParagraphOoxml(
      documentXml,
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    );
    expect(parsed[0].fieldTextSpans).toEqual([{ start: 11, end: 12 }]);
    const occurrence = extractFigureCitationOccurrences(
      'See Figure 2.',
      0,
      parsed[0].fieldTextSpans
    )[0];
    expect(occurrence.isFieldBacked).toBe(true);
  });

  it('recognizes SEQ Table fields as safe formatting-only table-number fields', () => {
    expect(isTableSequenceFieldInstruction(' SEQ Table \\* ARABIC ')).toBe(true);
    expect(isTableSequenceFieldInstruction(' REF _Ref123 ')).toBe(false);

    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
      <w:r><w:t>TABLE </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText> SEQ Table \\* ARABIC </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>1</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:t>. Results</w:t></w:r>
    </w:p></w:body></w:document>`;
    const parsed = parseParagraphOoxml(
      documentXml,
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    );

    expect(parsed[0].fieldTextSpans).toEqual([
      { start: 6, end: 7, kind: 'tableSequence' }
    ]);
    expect(validateRule('TABLE_PREFIX_BOLD_001', snapshot([
      paragraph({
        index: 0,
        classification: 'tableCaption',
        text: 'TABLE 1. Results',
        tablePrefixBold: false,
        hasField: true,
        fieldTextSpans: [{ start: 6, end: 7, kind: 'tableSequence' }]
      })
    ])).autoFixable).toBe(true);
  });

  it('evaluates line space above only for high-confidence inline figures', () => {
    const paragraphs = [
      paragraph({ index: 0, text: 'Body text above.', classification: 'body' }),
      paragraph({ index: 1, text: '', classification: 'other' }),
      paragraph({ index: 2, text: 'FIGURE 1. Results', classification: 'figureCaption' })
    ];
    const blocks = [
      { blockIndex: 0, kind: 'paragraph' as const, paragraphIndex: 0, text: paragraphs[0].text },
      { blockIndex: 1, kind: 'paragraph' as const, paragraphIndex: 1, drawingCount: 1, hasInlineDrawing: true, hasAnchoredDrawing: false },
      { blockIndex: 2, kind: 'paragraph' as const, paragraphIndex: 2, text: paragraphs[2].text }
    ];
    const doc = snapshot(paragraphs);
    doc.layout.documentBlocks = blocks;
    doc.crossReferences = analyzeCrossReferences(paragraphs, blocks);
    const missing = analyzeFigureSpaceAbove(doc);
    expect(missing.applicableFigureCount).toBe(1);
    expect(missing.issues[0].previousParagraphIndex).toBe(0);
    expect(validateRule('FIGURE_LINE_SPACE_ABOVE_001', doc).status).toBe('failed');

    const withBlank = snapshot([
      paragraph({ index: 0, text: 'Body text above.', classification: 'body' }),
      paragraph({ index: 1, text: '', classification: 'other', isBlank: true, blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      } }),
      paragraph({ index: 2, text: '', classification: 'other' }),
      paragraph({ index: 3, text: 'FIGURE 1. Results', classification: 'figureCaption' })
    ]);
    const blankBlocks = [
      { blockIndex: 0, kind: 'paragraph' as const, paragraphIndex: 0, text: withBlank.paragraphs[0].text },
      { blockIndex: 1, kind: 'paragraph' as const, paragraphIndex: 1 },
      { blockIndex: 2, kind: 'paragraph' as const, paragraphIndex: 2, drawingCount: 1, hasInlineDrawing: true, hasAnchoredDrawing: false },
      { blockIndex: 3, kind: 'paragraph' as const, paragraphIndex: 3, text: withBlank.paragraphs[3].text }
    ];
    withBlank.layout.documentBlocks = blankBlocks;
    withBlank.crossReferences = analyzeCrossReferences(withBlank.paragraphs, blankBlocks);
    expect(validateRule('FIGURE_LINE_SPACE_ABOVE_001', withBlank).status).toBe('passed');

    const multipleBlanks = snapshot([
      paragraph({ index: 0, text: 'Body text above.', classification: 'body' }),
      paragraph({ index: 1, text: '', classification: 'other', isBlank: true, blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      } }),
      paragraph({ index: 2, text: '', classification: 'other', isBlank: true, blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      } }),
      paragraph({ index: 3, text: '', classification: 'other' }),
      paragraph({ index: 4, text: 'FIGURE 1. Results', classification: 'figureCaption' })
    ]);
    const multipleBlankBlocks = [
      { blockIndex: 0, kind: 'paragraph' as const, paragraphIndex: 0, text: multipleBlanks.paragraphs[0].text },
      { blockIndex: 1, kind: 'paragraph' as const, paragraphIndex: 1 },
      { blockIndex: 2, kind: 'paragraph' as const, paragraphIndex: 2 },
      { blockIndex: 3, kind: 'paragraph' as const, paragraphIndex: 3, drawingCount: 1, hasInlineDrawing: true, hasAnchoredDrawing: false },
      { blockIndex: 4, kind: 'paragraph' as const, paragraphIndex: 4, text: multipleBlanks.paragraphs[4].text }
    ];
    multipleBlanks.layout.documentBlocks = multipleBlankBlocks;
    multipleBlanks.crossReferences = analyzeCrossReferences(multipleBlanks.paragraphs, multipleBlankBlocks);
    expect(validateRule('FIGURE_LINE_SPACE_ABOVE_001', multipleBlanks).status).toBe('failed');

    const anchored = snapshot(paragraphs);
    const anchoredBlocks = blocks.map(block => block.blockIndex === 1
      ? { ...block, hasInlineDrawing: false, hasAnchoredDrawing: true }
      : block);
    anchored.layout.documentBlocks = anchoredBlocks;
    anchored.crossReferences = analyzeCrossReferences(paragraphs, anchoredBlocks);
    expect(validateRule('FIGURE_LINE_SPACE_ABOVE_001', anchored).status).toBe('not_evaluated');

    const structuralBlank = snapshot([
      paragraph({ index: 0, text: 'Body text above.', classification: 'body' }),
      paragraph({ index: 1, text: '', classification: 'other', isBlank: true, blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: true,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: false
      } }),
      paragraph({ index: 2, text: '', classification: 'other' }),
      paragraph({ index: 3, text: 'FIGURE 1. Results', classification: 'figureCaption' })
    ]);
    structuralBlank.layout.documentBlocks = [
      { blockIndex: 0, kind: 'paragraph' as const, paragraphIndex: 0, text: structuralBlank.paragraphs[0].text },
      { blockIndex: 1, kind: 'paragraph' as const, paragraphIndex: 1 },
      { blockIndex: 2, kind: 'paragraph' as const, paragraphIndex: 2, drawingCount: 1, hasInlineDrawing: true, hasAnchoredDrawing: false },
      { blockIndex: 3, kind: 'paragraph' as const, paragraphIndex: 3, text: structuralBlank.paragraphs[3].text }
    ];
    structuralBlank.crossReferences = analyzeCrossReferences(structuralBlank.paragraphs, structuralBlank.layout.documentBlocks);
    expect(validateRule('FIGURE_LINE_SPACE_ABOVE_001', structuralBlank).status).toBe('not_evaluated');
  });

  it('keeps Reference Care in Submission Readiness instead of manuscript review', () => {
    const doc = snapshot([paragraph({ index: 0, text: '1. Reference', classification: 'reference' })]);
    doc.features = { ...ApplicabilityAnalyzer.analyze(doc), submissionMode: false };
    expect(validateRule('REFERENCE_CARE_001', doc).status).toBe('not_applicable');
    doc.features.submissionMode = true;
    expect(validateRule('REFERENCE_CARE_001', doc).status).toBe('checklist');
  });

  it('preserves ordered paragraph/table/drawing blocks for placement checks', () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>
      <w:p><w:r><w:t>TABLE 1. Caption</w:t></w:r></w:p>
      <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
      <w:p><w:r><w:drawing><wp:inline/></w:drawing></w:r></w:p>
      <w:p><w:r><w:t>FIGURE 1. Caption</w:t></w:r></w:p>
    </w:body></w:document>`;
    const blocks = parseDocumentBlocks(xml);
    expect(blocks.map(block => block.kind)).toEqual(['paragraph', 'table', 'paragraph', 'paragraph']);
    expect(blocks[0].paragraphIndex).toBe(0);
    expect(blocks[1].tableIndex).toBe(0);
    expect(blocks[2].drawingCount).toBe(1);
    expect(blocks[3].paragraphIndex).toBe(3);

    const analyzed = analyzeCrossReferences([
      paragraph({ index: 0, text: 'TABLE 1. Caption', classification: 'tableCaption' }),
      paragraph({ index: 1, text: '', classification: 'other' }),
      paragraph({ index: 2, text: '', classification: 'other' }),
      paragraph({ index: 3, text: 'FIGURE 1. Caption', classification: 'figureCaption' })
    ], blocks);
    expect(analyzed.tablePlacement[0].passed).toBe(true);
    expect(analyzed.tablePlacement[0].confidence).toBe('high');
    expect(analyzed.figurePlacement[0].passed).toBe(true);
    expect(analyzed.figurePlacement[0].confidence).toBe('high');
  });

  it('evaluates automated figure citation and placement rules from the shared model', () => {
    const paragraphs = [
      paragraph({ index: 0, text: 'The results are shown in Fig. 1.', classification: 'body' }),
      paragraph({ index: 1, text: '', classification: 'other' }),
      paragraph({ index: 2, text: 'FIGURE 1. Results', classification: 'figureCaption' })
    ];
    const doc = snapshot(paragraphs);
    doc.crossReferences = analyzeCrossReferences(paragraphs, [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, drawingCount: 0, text: paragraphs[0].text },
      { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1, drawingCount: 1, hasInlineDrawing: true, text: '' },
      { blockIndex: 2, kind: 'paragraph', paragraphIndex: 2, drawingCount: 0, text: paragraphs[2].text }
    ]);

    expect(validateRule('FIGURE_TEXT_CITATION_001', doc).status).toBe('passed');
    expect(validateRule('FIGURE_PRIOR_CITATION_001', doc).status).toBe('passed');
    expect(validateRule('FIGURE_CAPTION_POSITION_001', doc).status).toBe('passed');
  });

  it('does not make anchored figure placement high-confidence', () => {
    const paragraphs = [
      paragraph({ index: 0, text: '', classification: 'other' }),
      paragraph({ index: 1, text: 'FIGURE 1. Results', classification: 'figureCaption' })
    ];
    const doc = snapshot(paragraphs);
    doc.crossReferences = analyzeCrossReferences(paragraphs, [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, drawingCount: 1, hasAnchoredDrawing: true, text: '' },
      { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1, drawingCount: 0, text: paragraphs[1].text }
    ]);

    expect(doc.crossReferences.figurePlacement[0].passed).toBeNull();
    expect(doc.crossReferences.figurePlacement[0].confidence).toBe('low');
    expect(validateRule('FIGURE_CAPTION_POSITION_001', doc).status).toBe('not_evaluated');
  });

  it('does not guess when a figure caption is between two inline drawings', () => {
    const paragraphs = [
      paragraph({ index: 1, text: 'FIGURE 1. Results', classification: 'figureCaption' })
    ];
    const analyzed = analyzeCrossReferences(paragraphs, [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, drawingCount: 1, hasInlineDrawing: true, text: '' },
      { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1, drawingCount: 0, text: paragraphs[0].text },
      { blockIndex: 2, kind: 'paragraph', paragraphIndex: 2, drawingCount: 1, hasInlineDrawing: true, text: '' }
    ]);

    expect(analyzed.figurePlacement[0].passed).toBeNull();
    expect(analyzed.figurePlacement[0].confidence).toBe('low');
  });

  it('does not guess when a table caption is between two table objects', () => {
    const paragraphs = [
      paragraph({ index: 1, text: 'TABLE 1. Results', classification: 'tableCaption' })
    ];
    const analyzed = analyzeCrossReferences(paragraphs, [
      { blockIndex: 0, kind: 'table', tableIndex: 0, drawingCount: 0, text: 'A' },
      { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1, drawingCount: 0, text: paragraphs[0].text },
      { blockIndex: 2, kind: 'table', tableIndex: 1, drawingCount: 0, text: 'B' }
    ]);

    expect(analyzed.tablePlacement[0].passed).toBeNull();
    expect(analyzed.tablePlacement[0].confidence).toBe('low');
  });

  it('does not make mixed inline and anchored drawings high-confidence', () => {
    const paragraphs = [
      paragraph({ index: 1, text: 'FIGURE 1. Results', classification: 'figureCaption' })
    ];
    const analyzed = analyzeCrossReferences(paragraphs, [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, drawingCount: 2, hasInlineDrawing: true, hasAnchoredDrawing: true, text: '' },
      { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1, drawingCount: 0, text: paragraphs[0].text }
    ]);

    expect(analyzed.figurePlacement[0].passed).toBeNull();
    expect(analyzed.figurePlacement[0].confidence).toBe('low');
  });

  it('pairs consecutive figure objects with their respective captions', () => {
    const paragraphs = [
      paragraph({ index: 1, text: 'FIGURE 1. First', classification: 'figureCaption' }),
      paragraph({ index: 3, text: 'FIGURE 2. Second', classification: 'figureCaption' })
    ];
    const analyzed = analyzeCrossReferences(paragraphs, [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, drawingCount: 1, hasInlineDrawing: true, text: '' },
      { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1, drawingCount: 0, text: paragraphs[0].text },
      { blockIndex: 2, kind: 'paragraph', paragraphIndex: 2, drawingCount: 1, hasInlineDrawing: true, text: '' },
      { blockIndex: 3, kind: 'paragraph', paragraphIndex: 3, drawingCount: 0, text: paragraphs[1].text }
    ]);

    expect(analyzed.figurePlacement.map(item => item.passed)).toEqual([true, true]);
  });

  it('pairs consecutive table captions with the following tables', () => {
    const paragraphs = [
      paragraph({ index: 0, text: 'TABLE 1. First', classification: 'tableCaption' }),
      paragraph({ index: 2, text: 'TABLE 2. Second', classification: 'tableCaption' })
    ];
    const analyzed = analyzeCrossReferences(paragraphs, [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, text: paragraphs[0].text },
      { blockIndex: 1, kind: 'table', tableIndex: 0 },
      { blockIndex: 2, kind: 'paragraph', paragraphIndex: 2, text: paragraphs[1].text },
      { blockIndex: 3, kind: 'table', tableIndex: 1 }
    ]);

    expect(analyzed.tablePlacement.map(item => item.passed)).toEqual([true, true]);
  });

  it('checks figure prior citation against the figure object block', () => {
    const paragraphs = [
      paragraph({ index: 0, text: '', classification: 'other' }),
      paragraph({ index: 1, text: 'FIGURE 1. Results', classification: 'figureCaption' }),
      paragraph({ index: 2, text: 'As shown in Fig. 1.', classification: 'body' })
    ];
    const doc = snapshot(paragraphs);
    doc.crossReferences = analyzeCrossReferences(paragraphs, [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, drawingCount: 1, hasInlineDrawing: true, text: '' },
      { blockIndex: 1, kind: 'paragraph', paragraphIndex: 1, drawingCount: 0, text: paragraphs[1].text },
      { blockIndex: 2, kind: 'paragraph', paragraphIndex: 2, drawingCount: 0, text: paragraphs[2].text }
    ]);

    expect(doc.crossReferences.figurePriorCitationFailures.map(item => item.number)).toEqual([1]);
    expect(validateRule('FIGURE_PRIOR_CITATION_001', doc).status).toBe('failed');
  });

  it('reports an out-of-order table caption as a deterministic placement failure', () => {
    const paragraphs = [
      paragraph({ index: 0, text: 'TABLE 1. Results', classification: 'tableCaption' })
    ];
    const doc = snapshot(paragraphs);
    doc.crossReferences = analyzeCrossReferences(paragraphs, [
      { blockIndex: 0, kind: 'table', tableIndex: 0, drawingCount: 0, text: 'Cell' },
      { blockIndex: 1, kind: 'paragraph', paragraphIndex: 0, drawingCount: 0, text: paragraphs[0].text }
    ]);

    const finding = validateRule('TABLE_CAPTION_POSITION_001', doc);
    expect(finding.status).toBe('failed');
    expect(finding.paragraphIndices).toEqual([0]);
  });

  it('keeps consistency findings out of the compliance score', () => {
    const paragraphs = [
      paragraph({ index: 0, text: 'Prior work [9].', classification: 'body' }),
      paragraph({ index: 1, text: 'REFERENCES', classification: 'referencesHeading' }),
      paragraph({ index: 2, text: '1. Uncited reference', classification: 'reference' })
    ];
    const doc = snapshot(paragraphs);
    doc.crossReferences = analyzeCrossReferences(paragraphs, [], 1);
    const finding = validateRule('REFERENCE_TEXT_CITATION_COVERAGE_001', doc);
    expect(finding.assessmentRole).toBe('consistency');
    expect(finding.status).toBe('failed');
  });

  it('marks author superscript reviews not applicable when there is only one association', () => {
    const doc = snapshot([
      paragraph({ index: 0, classification: 'authorAffiliation', text: 'Faculty of Science' }),
      paragraph({ index: 1, classification: 'authorEmail', text: 'author@example.com' })
    ]);
    expect(validateRule('AUTHOR_AFFILIATION_SUPERSCRIPT_001', doc).status).toBe('not_applicable');
    expect(validateRule('AUTHOR_EMAIL_SUPERSCRIPT_001', doc).status).toBe('not_applicable');
  });

  it('keeps email-marker superscript validation applicable for a single marked email', () => {
    const doc = snapshot([
      paragraph({
        index: 0,
        classification: 'authorEmail',
        text: 'a) author@example.com',
        effectiveRuns: [
          { text: 'a)', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, verticalAlign: 'superscript', sources: {} },
          { text: ' author@example.com', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, verticalAlign: null, sources: {} }
        ]
      })
    ]);
    expect(validateRule('AUTHOR_EMAIL_SUPERSCRIPT_001', doc).status).toBe('passed');
  });

  it('validates affiliation and email association markers independently of italics', () => {
    const doc = snapshot([
      paragraph({
        index: 0,
        classification: 'authorAffiliation',
        text: '1 Faculty of Science',
        effectiveRuns: [
          { text: '1 ', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, verticalAlign: 'superscript', sources: {} },
          { text: 'Faculty of Science', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: true, allCaps: false, verticalAlign: null, sources: {} }
        ]
      }),
      paragraph({
        index: 1,
        classification: 'authorAffiliation',
        text: '2 Department of Physics',
        effectiveRuns: [
          { text: '2 ', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, verticalAlign: 'superscript', sources: {} },
          { text: 'Department of Physics', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: true, allCaps: false, verticalAlign: null, sources: {} }
        ]
      }),
      paragraph({
        index: 2,
        classification: 'authorEmail',
        text: 'a) author@example.com',
        effectiveRuns: [
          { text: 'a) ', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, verticalAlign: 'superscript', sources: {} },
          { text: 'author@example.com', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: true, allCaps: false, verticalAlign: null, sources: {} }
        ]
      }),
      paragraph({
        index: 3,
        classification: 'authorEmail',
        text: 'b) second@example.com',
        effectiveRuns: [
          { text: 'b) ', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, verticalAlign: 'superscript', sources: {} },
          { text: 'second@example.com', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: true, allCaps: false, verticalAlign: null, sources: {} }
        ]
      })
    ]);
    expect(validateRule('AUTHOR_AFFILIATION_SUPERSCRIPT_001', doc).status).toBe('passed');
    expect(validateRule('AUTHOR_EMAIL_SUPERSCRIPT_001', doc).status).toBe('passed');
  });

  it('does not require surrounding whitespace to be superscript', () => {
    const affiliationRuns = [
      { text: ' ', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, verticalAlign: null, sources: {} },
      { text: '2', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: false, allCaps: false, verticalAlign: 'superscript' as const, sources: {} },
      { text: 'Faculty of Science', fontName: 'Times New Roman', fontSize: 10, bold: false, italic: true, allCaps: false, verticalAlign: null, sources: {} }
    ];
    const doc = snapshot([
      paragraph({ index: 0, classification: 'authorAffiliation', text: ' 2Faculty of Science', effectiveRuns: affiliationRuns }),
      paragraph({ index: 1, classification: 'authorAffiliation', text: ' 2Department of Physics', effectiveRuns: affiliationRuns })
    ]);
    expect(validateRule('AUTHOR_AFFILIATION_SUPERSCRIPT_001', doc).status).toBe('passed');
  });

  it('does not mark a missing association marker as safely fixable', () => {
    const doc = snapshot([
      paragraph({ index: 0, classification: 'authorAffiliation', text: 'Faculty of Science', effectiveRuns: [] }),
      paragraph({ index: 1, classification: 'authorAffiliation', text: '2Department of Physics', effectiveRuns: [] })
    ]);
    const finding = validateRule('AUTHOR_AFFILIATION_SUPERSCRIPT_001', doc);
    expect(finding.status).toBe('failed');
    expect(finding.autoFixable).toBe(false);
  });

  it('keeps association review applicable when several authors share one paragraph', () => {
    const doc = snapshot([
      paragraph({ index: 0, classification: 'authorName', text: 'Ari1,a), Harin1,b), Nurudin1,c) and Popong2,d)' }),
      paragraph({ index: 1, classification: 'authorAffiliation', text: '1 Faculty of Science' }),
      paragraph({ index: 2, classification: 'authorAffiliation', text: '2 Department of Physics' }),
      paragraph({ index: 3, classification: 'authorEmail', text: 'a) ari@example.com' }),
      paragraph({ index: 4, classification: 'authorEmail', text: 'b) harin@example.com' })
    ]);
    expect(validateRule('AUTHOR_ASSOCIATION_001', doc).status).toBe('manual_review');
  });

  it('does not apply normal body first-line indentation to list paragraphs', () => {
    const finding = validateRule('BODY_FIRST_LINE_INDENT_001', snapshot([
      paragraph({ index: 0, text: '• First pattern', style: 'Paragraph (bulleted)', isListItem: true, firstLineIndent: -17.85 }),
      paragraph({ index: 1, text: 'Normal body paragraph', firstLineIndent: 14.4 })
    ]));
    expect(finding.status).toBe('passed');
    expect(finding.paragraphIndices).toEqual([]);
  });

  it('checks email superscript markers on logical lines within one Word paragraph', () => {
    const markerRun = (text: string) => ({
      text,
      fontName: 'Times New Roman',
      fontSize: 10,
      bold: false,
      italic: false,
      allCaps: false,
      verticalAlign: 'superscript' as const,
      sources: {}
    });
    const normalRun = (text: string) => ({
      text,
      fontName: 'Times New Roman',
      fontSize: 10,
      bold: false,
      italic: false,
      allCaps: false,
      verticalAlign: null,
      sources: {}
    });
    const doc = snapshot([
      paragraph({
        index: 0,
        classification: 'authorEmail',
        text: '2Omis Consulting Sdn. Bhd.\na) Corresponding author: wongtl@example.com\nb) ganzh@example.com',
        effectiveRuns: [
          normalRun('2Omis Consulting Sdn. Bhd.\n'),
          markerRun('a'),
          markerRun(') Corresponding author: wongtl@example.com\n'),
          markerRun('b'),
          markerRun(') ganzh@example.com')
        ]
      }),
      paragraph({
        index: 1,
        classification: 'authorEmail',
        text: 'c) leewp@example.com',
        effectiveRuns: [markerRun('c'), markerRun(') leewp@example.com')]
      })
    ]);
    expect(validateRule('AUTHOR_EMAIL_SUPERSCRIPT_001', doc).status).toBe('passed');
  });

  it('checks and fixes email body italics without including association markers', async () => {
    const run = (text: string, italic: boolean, verticalAlign: 'superscript' | null = null) => ({
      text,
      fontName: 'Times New Roman',
      fontSize: 10,
      bold: false,
      italic,
      allCaps: false,
      verticalAlign,
      sources: {}
    });
    const mixedText = '2Omis Consulting Sdn. Bhd.\na) Corresponding author: wongtl@example.com\nb) ganzh@example.com';
    const passingParagraph = paragraph({
      index: 0,
      classification: 'authorEmail',
      text: mixedText,
      effectiveRuns: [
        run('2Omis Consulting Sdn. Bhd.\n', false),
        run('a) ', false, 'superscript'),
        run('Corresponding author: wongtl@example.com\n', true),
        run('b) ', true, 'superscript'),
        run('ganzh@example.com', true)
      ]
    });
    expect(validateRule('AUTHOR_EMAIL_ITALIC_001', snapshot([passingParagraph])).status).toBe('passed');

    const failingParagraph = paragraph({
      ...passingParagraph,
      effectiveRuns: passingParagraph.effectiveRuns?.map(runValue =>
        runValue.text.includes('Corresponding') || runValue.text.includes('ganzh@example')
          ? { ...runValue, italic: false }
          : runValue
      )
    });
    const failure = validateRule('AUTHOR_EMAIL_ITALIC_001', snapshot([failingParagraph]));
    expect(failure.status).toBe('failed');
    expect(failure.detected).toContain('association markers are excluded');
    expect(failure.message).toContain('email text');

    const source = await readFile('src/word/WordDocumentAdapter.ts', 'utf8');
    expect(source).toContain('function authorItalicBodyTexts');
    expect(source).toContain('paragraphContentRange(entry.paragraph).search(bodyText');
  });

  it('prefers live email-body italic state over stale OOXML runs', () => {
    const finding = validateRule('AUTHOR_EMAIL_ITALIC_001', snapshot([
      paragraph({
        classification: 'authorEmail',
        text: 'a) author@example.com',
        italic: null,
        authorAssociationBodyItalic: true,
        effectiveRuns: [
          {
            text: 'a) ',
            fontName: 'Times New Roman',
            fontSize: 10,
            bold: false,
            italic: false,
            allCaps: false,
            verticalAlign: 'superscript',
            sources: {}
          },
          {
            text: 'author@example.com',
            fontName: 'Times New Roman',
            fontSize: 10,
            bold: false,
            italic: false,
            allCaps: false,
            verticalAlign: null,
            sources: {}
          }
        ]
      })
    ]));

    expect(finding.status).toBe('passed');
  });

  it('fails email italic validation when the live body range is not italic', () => {
    const finding = validateRule('AUTHOR_EMAIL_ITALIC_001', snapshot([
      paragraph({
        classification: 'authorEmail',
        text: 'a) author@example.com',
        italic: null,
        authorAssociationBodyItalic: false,
        effectiveRuns: [
          {
            text: 'a) author@example.com',
            fontName: 'Times New Roman',
            fontSize: 10,
            bold: false,
            italic: true,
            allCaps: false,
            verticalAlign: null,
            sources: {}
          }
        ]
      })
    ]));

    expect(finding.status).toBe('failed');
  });

  it('falls back to OOXML runs when live author-body italic state is unavailable', () => {
    const finding = validateRule('AUTHOR_EMAIL_ITALIC_001', snapshot([
      paragraph({
        classification: 'authorEmail',
        text: 'a) author@example.com',
        italic: null,
        authorAssociationBodyItalic: null,
        effectiveRuns: [
          {
            text: 'a) ',
            fontName: 'Times New Roman',
            fontSize: 10,
            bold: false,
            italic: false,
            allCaps: false,
            verticalAlign: 'superscript',
            sources: {}
          },
          {
            text: 'author@example.com',
            fontName: 'Times New Roman',
            fontSize: 10,
            bold: false,
            italic: true,
            allCaps: false,
            verticalAlign: null,
            sources: {}
          }
        ]
      })
    ]));

    expect(finding.status).toBe('passed');
  });

  it('uses live affiliation-body italic state while excluding its marker', () => {
    const finding = validateRule('AUTHOR_AFFILIATION_ITALIC_001', snapshot([
      paragraph({
        classification: 'authorAffiliation',
        text: '1 University Name',
        italic: null,
        authorAssociationBodyItalic: true,
        effectiveRuns: [{
          text: '1 University Name',
          fontName: 'Times New Roman',
          fontSize: 10,
          bold: false,
          italic: false,
          allCaps: false,
          verticalAlign: null,
          sources: {}
        }]
      })
    ]));

    expect(finding.status).toBe('passed');
  });

  it('uses paragraph content ranges for text replacement corrections', async () => {
    const source = await readFile('src/word/WordDocumentAdapter.ts', 'utf8');
    expect(source).toContain('function paragraphContentRange(paragraph: Word.Paragraph)');
    expect(source).toContain('paragraph.getRange(Word.RangeLocation.content)');
    expect(source).toContain('content.font.name = String(expected)');
    expect(source).toContain('static async applyParagraphTypography');
    expect(source).not.toContain('p.getRange().insertText(revised');
    expect(source).not.toContain('x.paragraph.getRange().insertText(revised');
  });

  it('activates submission checklist rules only in Submission Readiness mode', () => {
    const doc = snapshot([]);
    expect(validateRule('PDF_FONT_EMBED_001', doc).status).toBe('not_applicable');
    doc.features = {
      authorCount: 0, affiliationCount: 0, emailCount: 0, hasMultipleAuthors: false,
      hasMultipleAffiliations: false, hasMultipleEmails: false, hasAuthorAssociationMarkers: false,
      hasEmailAssociationMarkers: false,
      hasDuplicateAffiliationCandidates: false, hasMultipleAssociations: false, figureCount: 0,
      tableCount: 0, figureCaptionCount: 0, tableCaptionCount: 0, figureObjectCount: 0, tableObjectCount: 0,
      equationCount: 0, referenceCount: 0, hasFigures: false, hasTables: false,
      hasFigureCaptions: false, hasTableCaptions: false, hasFigureObjects: false, hasTableObjects: false,
      hasEquations: false, hasReferences: false, hasMultiPartFigureCandidates: false,
      hasFloatingFigures: false, hasPdf: false, submissionMode: true
    };
    expect(validateRule('PDF_FONT_EMBED_001', doc).status).toBe('checklist');
  });

  it('keeps submission figure reviews out of ordinary manuscript validation', () => {
    const doc = snapshot([
      paragraph({ index: 0, classification: 'figureCaption', text: 'FIGURE 1. Results' })
    ]);
    expect(validateRule('FIGURE_READABILITY_001', doc).status).toBe('not_applicable');
    doc.features = {
      authorCount: 0, affiliationCount: 0, emailCount: 0, hasMultipleAuthors: false,
      hasMultipleAffiliations: false, hasMultipleEmails: false, hasEmailAssociationMarkers: false,
      hasAuthorAssociationMarkers: false, hasDuplicateAffiliationCandidates: false, hasMultipleAssociations: false,
      figureCount: 1, tableCount: 0, figureCaptionCount: 1, tableCaptionCount: 0, figureObjectCount: 1, tableObjectCount: 0,
      equationCount: 0, referenceCount: 0, hasFigures: true, hasTables: false,
      hasFigureCaptions: true, hasTableCaptions: false, hasFigureObjects: true, hasTableObjects: false,
      hasEquations: false, hasReferences: false, hasMultiPartFigureCandidates: false,
      hasFloatingFigures: false, hasPdf: false, submissionMode: true
    };
    expect(validateRule('FIGURE_READABILITY_001', doc).status).toBe('checklist');
    expect(validateRule('COPYRIGHT_PERMISSION_001', doc).status).toBe('checklist');
  });

  it('detects missing equation center/right tabs', () => {
    const finding = validateRule('EQUATION_TABS_001', snapshot([
      paragraph({ index: 0, classification: 'equation', text: 'x = 1 (1)', tabStops: [] })
    ]));
    expect(finding.status).toBe('failed');
  });

  it('evaluates table top, bottom, and header rules from ordered OOXML blocks', () => {
    const paragraphs = [paragraph({ index: 0, classification: 'tableCaption', text: 'TABLE 1. Results' })];
    const doc = snapshot(paragraphs);
    doc.layout.documentBlocks = [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, text: paragraphs[0].text },
      {
        blockIndex: 1,
        kind: 'table',
        tableIndex: 0,
        tableTopBorder: true,
        tableBottomBorder: false,
        tableHeaderRule: true,
        tableHeaderRowCount: 1,
        tableVerticalRules: true,
        tableExtraHorizontalRules: true
      }
    ];
    doc.crossReferences = {
      referenceDefinitions: [], referenceCitations: [], referenceCitationStyle: 'unknown',
      figureDefinitions: [], figureCitations: [], tableDefinitions: [{ kind: 'table', number: 1, paragraphIndex: 0 }],
      tableCitations: [], referenceCoverageMissing: [], invalidReferenceCitations: [], referenceSequenceViolations: [],
      figureCoverageMissing: [], tableCoverageMissing: [], figurePlacement: [],
      tablePlacement: [{ kind: 'table', number: 1, captionParagraphIndex: 0, objectBlockIndex: 1, passed: true, confidence: 'high' }],
      figurePriorCitationFailures: [], figurePriorCitationUnknown: []
    };
    expect(validateRule('TABLE_TOP_RULE_001', doc).status).toBe('passed');
    expect(validateRule('TABLE_BOTTOM_RULE_001', doc).status).toBe('failed');
    expect(validateRule('TABLE_HEADER_RULE_001', doc).status).toBe('passed');
    expect(validateRule('TABLE_NO_VERTICAL_RULES_001', doc).status).toBe('failed');
    expect(validateRule('TABLE_NO_EXTRA_HORIZONTAL_RULES_001', doc).status).toBe('failed');
  });

  it('resolves table outer rules from cell-level borders and excludes the final bottom rule from extras', () => {
    const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:tbl>
        <w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="single"/><w:bottom w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:tcBorders><w:top w:val="single"/><w:bottom w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc></w:tr>
        <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
        <w:tr><w:tc><w:tcPr><w:tcBorders><w:bottom w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:tcBorders><w:bottom w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc></w:tr>
      </w:tbl>
    </w:body></w:document>`;
    const blocks = parseDocumentBlocks(documentXml);
    expect(blocks).toHaveLength(1);
    expect(blocks[0].tableTopBorder).toBe(true);
    expect(blocks[0].tableBottomBorder).toBe(true);
    expect(blocks[0].tableHeaderRule).toBe(true);
    expect(blocks[0].tableVerticalRules).toBe(false);
    expect(blocks[0].tableExtraHorizontalRules).toBe(false);
  });

  it('requires a complete header separator and accepts the lower-row top-border representation', () => {
    const complete = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
      <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
      <w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:tcBorders><w:top w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc></w:tr>
    </w:tbl></w:body></w:document>`;
    expect(parseDocumentBlocks(complete)[0].tableHeaderRule).toBe(true);

    const partial = complete.replace(
      '<w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:tcBorders><w:top w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc></w:tr>',
      '<w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="single"/></w:tcBorders></w:tcPr><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>'
    );
    expect(parseDocumentBlocks(partial)[0].tableHeaderRule).toBe(false);

    const disabledHeader = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:trPr><w:tblHeader w:val="0"/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>`;
    expect(parseDocumentBlocks(disabledHeader)[0].tableHeaderRule).toBeNull();
  });

  it('carries the detected multi-row header count into table corrections', () => {
    const paragraphs = [paragraph({ index: 0, classification: 'tableCaption', text: 'TABLE 1. Results' })];
    const doc = snapshot(paragraphs);
    doc.layout.documentBlocks = [
      { blockIndex: 0, kind: 'paragraph', paragraphIndex: 0, text: paragraphs[0].text },
      { blockIndex: 1, kind: 'table', tableIndex: 0, tableHeaderRule: false, tableHeaderRowCount: 2 }
    ];
    doc.crossReferences = {
      referenceDefinitions: [], referenceCitations: [], referenceCitationStyle: 'unknown',
      figureDefinitions: [], figureCitations: [], tableDefinitions: [{ kind: 'table', number: 1, paragraphIndex: 0 }],
      tableCitations: [], referenceCoverageMissing: [], invalidReferenceCitations: [], referenceSequenceViolations: [],
      figureCoverageMissing: [], tableCoverageMissing: [], figurePlacement: [],
      tablePlacement: [{ kind: 'table', number: 1, captionParagraphIndex: 0, objectBlockIndex: 1, passed: true, confidence: 'high' }],
      figurePriorCitationFailures: [], figurePriorCitationUnknown: []
    };
    const finding = validateRule('TABLE_HEADER_RULE_001', doc);
    expect(finding.status).toBe('failed');
    expect(finding.fixTableIndices).toEqual([0]);
    expect(finding.fixTableHeaderRowCounts).toEqual([2]);
  });

  it('keeps submission-only checks out of ordinary manuscript review', () => {
    const finding = validateRule('PDF_FONT_EMBED_001', snapshot([]));
    expect(finding.status).toBe('not_applicable');
  });

  it('shows failed consistency findings in Failed while keeping them separate in scoring', () => {
    const complianceFailure = correctionFinding({
      ruleId: 'FIGURE_TEXT_CITATION_001',
      assessmentRole: 'compliance',
      status: 'failed'
    });
    const consistencyFailure = correctionFinding({
      ruleId: 'TABLE_TEXT_CITATION_001',
      assessmentRole: 'consistency',
      status: 'failed'
    });
    const consistencyPass = correctionFinding({
      ruleId: 'TEXT_MULTIPLE_SPACES_001',
      assessmentRole: 'consistency',
      status: 'passed'
    });
    const consistencyNotEvaluated = correctionFinding({
      ruleId: 'AFFILIATION_EMAIL_BLANK_001',
      assessmentRole: 'consistency',
      status: 'not_evaluated'
    });
    const notApplicable = correctionFinding({
      ruleId: 'ABSTRACT_SPACING_001',
      assessmentRole: 'compliance',
      status: 'not_applicable'
    });
    const findings = [complianceFailure, consistencyFailure, consistencyPass, consistencyNotEvaluated, notApplicable];

    expect(filterFindings(findings, 'failed').map(finding => finding.ruleId))
      .toEqual(['FIGURE_TEXT_CITATION_001', 'TABLE_TEXT_CITATION_001']);
    expect(filterFindings(findings, 'consistency').map(finding => finding.ruleId))
      .toEqual(['TABLE_TEXT_CITATION_001', 'TEXT_MULTIPLE_SPACES_001', 'AFFILIATION_EMAIL_BLANK_001']);
    expect(filterFindings(findings, 'passed').map(finding => finding.ruleId))
      .toEqual(['TEXT_MULTIPLE_SPACES_001']);
    expect(filterFindings(findings, 'all')).toHaveLength(4);
  });

  it('merges overlapping typography findings into one adapter operation', async () => {
    const applyTypography = vi
      .spyOn(WordDocumentAdapter, 'applyParagraphTypography')
      .mockResolvedValue(undefined);

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'BODY_FONT_001',
          paragraphIndices: [4],
          property: 'font.name',
          rawExpected: 'Times New Roman'
        }),
        correctionFinding({
          ruleId: 'BODY_SIZE_001',
          ruleName: 'Body size',
          paragraphIndices: [4],
          property: 'font.size',
          rawExpected: 10,
          expected: '10 pt'
        })
      ]);

      expect(applyTypography).toHaveBeenCalledTimes(1);
      expect(applyTypography).toHaveBeenCalledWith(
        [4],
        { fontName: 'Times New Roman', fontSize: 10 }
      );
      expect(result.requestedFixes).toBe(2);
      expect(result.executedFixes).toBe(2);
      expect(result.executionFailures).toBe(0);
      expect(result.attempts.every(attempt => attempt.executionSucceeded)).toBe(true);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('formats the complete reference list and does not rebuild numbering for style-only fixes', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const referenceDocument = snapshot([
      paragraph({ index: 1, classification: 'reference', text: '1. First reference' }),
      paragraph({ index: 2, classification: 'reference', text: '2. Second reference' }),
      paragraph({ index: 3, classification: 'reference', text: '3. Third reference' })
    ]);
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(referenceDocument);
    const applyReference = vi.spyOn(WordDocumentAdapter, 'applyReferenceFormatting').mockResolvedValue(undefined);
    (globalThis as { Word?: unknown }).Word = { run: vi.fn() };

    try {
      const styleResult = await CorrectionService.apply([correctionFinding({
        ruleId: 'REFERENCE_FONT_001',
        ruleName: 'Reference style font',
        target: 'reference',
        paragraphIndices: [2],
        property: 'reference.style.fontName',
        rawExpected: 'Times New Roman'
      })]);
      expect(applyReference).toHaveBeenNthCalledWith(1, [1, 2, 3], {
        applyNumbering: false,
        properties: ['reference.style.fontName']
      });
      expect(styleResult.attempts[0].paragraphIndices).toEqual([1, 2, 3]);

      await CorrectionService.apply([correctionFinding({
        ruleId: 'REFERENCES_NUMBERING_001',
        ruleName: 'Reference numbering',
        target: 'reference',
        paragraphIndices: [2],
        fixParagraphIndices: [2],
        property: 'text.referenceNumbering',
        rawExpected: true
      })]);
      expect(applyReference).toHaveBeenNthCalledWith(2, [1, 2, 3], {
        applyNumbering: true,
        properties: ['text.referenceNumbering']
      });
      expect(snapshotSpy).toHaveBeenCalled();
    } finally {
      vi.restoreAllMocks();
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('distinguishes a completed reference operation from failed in-service verification', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const invalidReference = paragraph({
      index: 1,
      classification: 'reference',
      style: 'Reference',
      referenceStyle: {
        name: 'Reference',
        type: 'paragraph',
        basedOn: 'Paragraph',
        nextParagraphStyle: 'Reference',
        fontName: 'Calibri',
        fontSize: 10,
        alignment: 'justified',
        leftIndent: 0,
        hangingIndent: 21.6
      }
    });
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(snapshot([invalidReference]));
    const applyReference = vi.spyOn(WordDocumentAdapter, 'applyReferenceFormatting').mockResolvedValue(undefined);
    (globalThis as { Word?: unknown }).Word = { run: vi.fn() };

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'REFERENCE_FONT_001',
          ruleName: 'Reference style font',
          target: 'reference',
          paragraphIndices: [1],
          property: 'reference.style.fontName',
          rawExpected: 'Times New Roman'
        })
      ], strategy);

      expect(applyReference).toHaveBeenCalledWith([1], {
        applyNumbering: false,
        properties: ['reference.style.fontName']
      });
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(1);
      expect(result.attempts[0].executionSucceeded).toBe(true);
      expect(result.attempts[0].verificationSucceeded).toBe(false);
      expect(result.attempts[0].message).toContain('not explicitly passed');
      expect(snapshotSpy).toHaveBeenCalledTimes(2);
    } finally {
      vi.restoreAllMocks();
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('fails reference verification when the post-correction reference evidence disappears', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const before = snapshot([paragraph({
      index: 1,
      classification: 'reference',
      style: 'Reference',
      text: 'Author, Example (2025).',
      referenceStyle: {
        name: 'Reference', type: 'paragraph', basedOn: 'Paragraph', nextParagraphStyle: 'Reference',
        fontName: 'Times New Roman', fontSize: 10, alignment: 'justified', leftIndent: 0, hangingIndent: 21.6
      }
    })]);
    const after = snapshot([]);
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot')
      .mockResolvedValueOnce(before)
      .mockResolvedValueOnce(after);
    const applyReference = vi.spyOn(WordDocumentAdapter, 'applyReferenceFormatting').mockResolvedValue(undefined);
    (globalThis as { Word?: unknown }).Word = { run: vi.fn() };

    try {
      const result = await CorrectionService.apply([correctionFinding({
        ruleId: 'REFERENCE_FONT_001',
        ruleName: 'Reference style font',
        target: 'reference',
        paragraphIndices: [1],
        property: 'reference.style.fontName',
        rawExpected: 'Times New Roman'
      })], strategy);

      expect(applyReference).toHaveBeenCalledWith([1], {
        applyNumbering: false,
        properties: ['reference.style.fontName']
      });
      expect(result.attempts[0].executionSucceeded).toBe(true);
      expect(result.attempts[0].verificationSucceeded).toBe(false);
      expect(result.executionFailures).toBe(1);
      expect(result.attempts[0].message).toContain('reference count or text changed');
      expect(snapshotSpy).toHaveBeenCalledTimes(2);
    } finally {
      vi.restoreAllMocks();
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('withholds non-sequential manual reference conversion before deleting prefixes', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const previousOffice = (globalThis as { Office?: unknown }).Office;
    const refs = [1, 3, 2].map(number => ({
      text: `${number}. Reference ${number}`,
      isListItem: false,
      load: vi.fn(),
      getOoxml: vi.fn(() => ({ value: '<w:p><w:pPr/></w:p>' }))
    }));
    const context = {
      document: {
        body: { paragraphs: { items: refs, load: vi.fn() } },
        styles: {
          getByNameOrNullObject: vi.fn((name: string) => ({
            isNullObject: false,
            load: vi.fn(),
            name
          }))
        }
      },
      sync: vi.fn().mockResolvedValue(undefined)
    };
    (globalThis as { Office?: unknown }).Office = {
      context: { requirements: { isSetSupported: vi.fn(() => true) } }
    };
    (globalThis as { Word?: unknown }).Word = {
      run: async (callback: (value: typeof context) => Promise<void>) => callback(context)
    };

    try {
      await expect(WordDocumentAdapter.applyReferenceFormatting([0, 1, 2]))
        .rejects.toThrow('not sequential');
      expect(refs.every(ref => ref.load.mock.calls.length === 1)).toBe(true);
    } finally {
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
    }
  });

  it('withholds non-sequential automatic reference conversion before rebuilding the list', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const previousOffice = (globalThis as { Office?: unknown }).Office;
    const refs = [1, 3, 2].map(number => ({
      text: `Reference ${number}`,
      isListItem: true,
      listItem: { listString: `${number}.`, load: vi.fn() },
      load: vi.fn(),
      getOoxml: vi.fn(() => ({ value: '<w:p><w:pPr/></w:p>' }))
    }));
    const context = {
      document: {
        body: { paragraphs: { items: refs, load: vi.fn() } },
        styles: {
          getByNameOrNullObject: vi.fn((name: string) => ({
            isNullObject: false,
            load: vi.fn(),
            name
          }))
        }
      },
      sync: vi.fn().mockResolvedValue(undefined)
    };
    (globalThis as { Office?: unknown }).Office = {
      context: { requirements: { isSetSupported: vi.fn(() => true) } }
    };
    (globalThis as { Word?: unknown }).Word = {
      run: async (callback: (value: typeof context) => Promise<void>) => callback(context)
    };

    try {
      await expect(WordDocumentAdapter.applyReferenceFormatting([0, 1, 2]))
        .rejects.toThrow('automatic reference numbers are not sequential');
      expect(refs.every(ref => ref.listItem.load.mock.calls.length === 1)).toBe(true);
    } finally {
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
    }
  });

  it('allows style-only reference correction without the desktop list API', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const previousOffice = (globalThis as { Office?: unknown }).Office;
    const contentRange = { font: {} };
    const reference = {
      text: 'Author, Reference (2025).',
      isListItem: false,
      load: vi.fn(),
      getOoxml: vi.fn(() => ({ value: '<w:p><w:pPr><w:ind w:left="100" w:hanging="50"/></w:pPr><w:r><w:t>Author</w:t></w:r></w:p>' })),
      insertOoxml: vi.fn(),
      getRange: vi.fn(() => contentRange)
    };
    const referenceStyle = {
      isNullObject: false,
      type: 'paragraph',
      nameLocal: 'Reference',
      load: vi.fn(),
      font: { name: 'Times New Roman', size: 10 },
      paragraphFormat: {
        alignment: 'Justified',
        leftIndent: 0,
        firstLineIndent: -21.6
      }
    };
    const paragraphStyle = { isNullObject: false, load: vi.fn() };
    const context = {
      document: {
        body: { paragraphs: { items: [reference], load: vi.fn() } },
        styles: {
          getByNameOrNullObject: vi.fn((name: string) =>
            name === 'Reference' ? referenceStyle : paragraphStyle)
        }
      },
      sync: vi.fn().mockResolvedValue(undefined)
    };
    (globalThis as { Office?: unknown }).Office = {
      context: {
        requirements: {
          isSetSupported: vi.fn((setName: string) => setName !== 'WordApiDesktop')
        }
      }
    };
    (globalThis as { Word?: unknown }).Word = {
      Alignment: { justified: 'Justified' },
      RangeLocation: { content: 'Content' },
      InsertLocation: { replace: 'Replace' },
      run: async (callback: (value: typeof context) => Promise<void>) => callback(context)
    };

    try {
      await expect(WordDocumentAdapter.applyReferenceFormatting([0], { applyNumbering: false }))
        .resolves.toBeUndefined();
      expect(reference.getOoxml).toHaveBeenCalled();
      expect(reference.insertOoxml).toHaveBeenCalledWith(
        expect.not.stringContaining('<w:ind'),
        'Replace'
      );
    } finally {
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
    }
  });

  it('repairs a direct reference indent without rewriting a compliant raw AIP style', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const previousOffice = (globalThis as { Office?: unknown }).Office;
    const reference = {
      text: 'Author, Reference (2025).',
      style: 'Reference',
      isListItem: false,
      load: vi.fn(),
      getOoxml: vi.fn(() => ({
        value: '<w:p><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:t>Author</w:t></w:r></w:p>'
      })),
      insertOoxml: vi.fn()
    };
    const referenceStyle = {
      isNullObject: false,
      type: 'paragraph',
      nameLocal: 'Reference',
      load: vi.fn(),
      baseStyle: 'Paragraph',
      nextParagraphStyle: 'Reference',
      font: { name: 'Times New Roman', size: 10 },
      paragraphFormat: { leftIndent: 21.3, firstLineIndent: -21.3 }
    };
    const paragraphStyle = { isNullObject: false, load: vi.fn() };
    const context = {
      document: {
        body: { paragraphs: { items: [reference], load: vi.fn() } },
        styles: {
          getByNameOrNullObject: vi.fn((name: string) =>
            name === 'Reference' ? referenceStyle : paragraphStyle)
        }
      },
      sync: vi.fn().mockResolvedValue(undefined)
    };
    (globalThis as { Office?: unknown }).Office = {
      context: { requirements: { isSetSupported: vi.fn(() => true) } }
    };
    (globalThis as { Word?: unknown }).Word = {
      InsertLocation: { replace: 'Replace' },
      run: async (callback: (value: typeof context) => Promise<void>) => callback(context)
    };

    try {
      await WordDocumentAdapter.applyReferenceFormatting([0], {
        applyNumbering: false,
        properties: ['reference.style.leftIndent']
      });
      expect(referenceStyle.paragraphFormat.leftIndent).toBe(21.3);
      expect(referenceStyle.paragraphFormat.firstLineIndent).toBe(-21.3);
      expect(reference.insertOoxml).toHaveBeenCalledWith(
        expect.not.stringContaining('w:left="720"'),
        'Replace'
      );

      referenceStyle.paragraphFormat.leftIndent = 30;
      referenceStyle.paragraphFormat.firstLineIndent = -21.3;
      await WordDocumentAdapter.applyReferenceFormatting([0], {
        applyNumbering: false,
        properties: ['reference.style.leftIndent']
      });
      expect(referenceStyle.paragraphFormat.leftIndent).toBe(21.3);
      expect(referenceStyle.paragraphFormat.firstLineIndent).toBe(-21.3);

      for (const firstLineIndent of [-10, -30]) {
        referenceStyle.paragraphFormat.leftIndent = 30;
        referenceStyle.paragraphFormat.firstLineIndent = firstLineIndent;
        await WordDocumentAdapter.applyReferenceFormatting([0], {
          applyNumbering: false,
          properties: ['reference.style.hangingIndent']
        });
        expect(referenceStyle.paragraphFormat.leftIndent).toBe(21.3);
        expect(referenceStyle.paragraphFormat.firstLineIndent).toBe(-21.3);
      }

      delete (referenceStyle.paragraphFormat as { firstLineIndent?: number }).firstLineIndent;
      await expect(WordDocumentAdapter.applyReferenceFormatting([0], {
        applyNumbering: false,
        properties: ['reference.style.leftIndent']
      })).rejects.toThrow('representation could not be determined');
      await expect(WordDocumentAdapter.applyReferenceFormatting([0], {
        applyNumbering: false,
        properties: ['reference.style.hangingIndent']
      })).rejects.toThrow('representation could not be determined');
    } finally {
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
    }
  });

  it('withholds numbering when a gap paragraph contains a structural page break', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const previousOffice = (globalThis as { Office?: unknown }).Office;
    const inspect = vi.spyOn(DocxInspector, 'inspectCurrentDocument').mockResolvedValue({
      ooxmlAvailable: true,
      paragraphOoxml: [{
        index: 1,
        blankParagraph: {
          whitespaceOnly: true,
          hasManualBreak: true,
          hasExplicitPageBreak: true,
          hasLastRenderedPageBreak: false,
          hasSectionBreak: false,
          hasDrawing: false,
          hasField: false,
          hasEquation: false,
          hasNumbering: false,
          hasBookmarkContent: false,
          safeToDelete: false
        },
        layoutSignals: {
          hasDrawing: false,
          hasAnchoredDrawing: false,
          hasCustomTabs: false,
          hasNumbering: false
        }
      }]
    });
    const run = vi.fn();
    (globalThis as { Office?: unknown }).Office = {
      context: { requirements: { isSetSupported: vi.fn(() => true) } }
    };
    (globalThis as { Word?: unknown }).Word = { run };

    try {
      await expect(WordDocumentAdapter.applyReferenceFormatting([0, 2]))
        .rejects.toThrow('structural paragraph');
      expect(inspect).toHaveBeenCalledTimes(1);
      expect(run).not.toHaveBeenCalled();
    } finally {
      inspect.mockRestore();
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
      if (previousOffice === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previousOffice;
    }
  });

  it('fails grouped page-size verification when refreshed section metadata disappears', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const applyLetter = vi.spyOn(WordDocumentAdapter, 'applyLetterPageSize').mockResolvedValue(undefined);
    const missingSectionMetadata = snapshot([]);
    missingSectionMetadata.layout.sections = [];
    vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(missingSectionMetadata);
    (globalThis as { Word?: unknown }).Word = { run: vi.fn() };

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'PAGE_WIDTH_001',
          ruleName: 'Page width',
          target: 'document',
          paragraphIndices: [],
          property: 'layout.pageWidthInches',
          rawExpected: 8.5
        }),
        correctionFinding({
          ruleId: 'PAGE_HEIGHT_001',
          ruleName: 'Page height',
          target: 'document',
          paragraphIndices: [],
          property: 'layout.pageHeightInches',
          rawExpected: 11
        })
      ]);

      expect(applyLetter).toHaveBeenCalledTimes(1);
      expect(result.executionFailures).toBe(2);
      expect(result.attempts.every(attempt => attempt.executionSucceeded === false)).toBe(true);
      expect(result.attempts[0].message).toContain('metadata was unavailable');
    } finally {
      vi.restoreAllMocks();
      if (previousWord === undefined) delete (globalThis as { Word?: unknown }).Word;
      else (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('routes author-name size through the marker-preserving correction path', async () => {
    const applyAuthorName = vi
      .spyOn(WordDocumentAdapter, 'applyAuthorNameTypographyPreservingMarkers')
      .mockResolvedValue(undefined);

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'AUTHOR_NAME_SIZE_001',
          ruleName: 'Author name size',
          target: 'authorName',
          paragraphIndices: [2],
          property: 'font.size',
          rawExpected: 14,
          expected: '14 pt'
        })
      ]);

      expect(applyAuthorName).toHaveBeenCalledWith([2], { fontSize: 14 });
      expect(result.requestedFixes).toBe(1);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('routes Roman table labels through the prefix-only table correction path', async () => {
    const normalizeTable = vi
      .spyOn(WordDocumentAdapter, 'normalizeTableCaptionPrefixes')
      .mockResolvedValue(undefined);

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'TABLE_CAPTION_PREFIX_001',
          ruleName: 'Table caption prefix',
          target: 'tableCaption',
          paragraphIndices: [3],
          fixParagraphIndices: [3],
          property: 'text.tablePrefix',
          rawExpected: true,
          expected: 'TABLE n.'
        })
      ]);

      expect(normalizeTable).toHaveBeenCalledWith([3]);
      expect(result.requestedFixes).toBe(1);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('refuses a stale body typography fix when the fresh paragraph is layout-sensitive', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const applyTypography = vi.spyOn(WordDocumentAdapter, 'applyParagraphTypography').mockResolvedValue(undefined);
    vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(snapshot([
      paragraph({ index: 4, layoutSensitive: true })
    ]));
    (globalThis as { Word?: unknown }).Word = { run: vi.fn() };

    try {
      const result = await CorrectionService.apply([
        correctionFinding({ paragraphIndices: [4] })
      ]);
      expect(applyTypography).not.toHaveBeenCalled();
      expect(result.executionFailures).toBe(1);
      expect(result.attempts[0].message).toContain('protected positioned/layout region');
    } finally {
      vi.restoreAllMocks();
      (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('protects body typography without filtering an explicitly selected caption fix', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const applyTypography = vi.spyOn(WordDocumentAdapter, 'applyParagraphTypography').mockResolvedValue(undefined);
    vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(snapshot([
      paragraph({ index: 4, layoutSensitive: true, classification: 'other' }),
      paragraph({ index: 5, classification: 'tableCaption' }),
      paragraph({ index: 6, layoutSensitive: true, classification: 'body' })
    ]));
    (globalThis as { Word?: unknown }).Word = { run: vi.fn() };

    try {
      const result = await CorrectionService.apply([
        correctionFinding({ ruleId: 'BODY_SIZE_001', paragraphIndices: [4], property: 'font.size', rawExpected: 10 }),
        correctionFinding({ ruleId: 'TABLE_CAPTION_SIZE_001', target: 'tableCaption', paragraphIndices: [5], property: 'font.size', rawExpected: 10 }),
        correctionFinding({ ruleId: 'CONCLUSION_BODY_FONT_001', target: 'document', paragraphIndices: [6], property: 'document.conclusionTypography', rawExpected: true })
      ]);
      expect(applyTypography).toHaveBeenCalledWith([5], { fontSize: 10 });
      expect(result.attempts.find(attempt => attempt.ruleId === 'BODY_SIZE_001')?.executionSucceeded).toBe(false);
      expect(result.attempts.find(attempt => attempt.ruleId === 'TABLE_CAPTION_SIZE_001')?.executionSucceeded).toBe(true);
      expect(result.attempts.find(attempt => attempt.ruleId === 'CONCLUSION_BODY_FONT_001')?.executionSucceeded).toBe(false);
    } finally {
      vi.restoreAllMocks();
      (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('fails closed when the paragraph count mapping is unreliable during correction', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const applyTypography = vi.spyOn(WordDocumentAdapter, 'applyParagraphTypography').mockResolvedValue(undefined);
    const doc = snapshot([paragraph({ index: 4 })]);
    doc.layout.paragraphMappingAligned = false;
    vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(doc);
    (globalThis as { Word?: unknown }).Word = { run: vi.fn() };

    try {
      const result = await CorrectionService.apply([
        correctionFinding({ paragraphIndices: [4] })
      ]);
      expect(applyTypography).not.toHaveBeenCalled();
      expect(result.executionFailures).toBe(1);
      expect(result.attempts[0].message).toContain('layout could not be inspected safely');
    } finally {
      vi.restoreAllMocks();
      (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('fails closed for an individual live/OOXML paragraph text mismatch', async () => {
    const previousWord = (globalThis as { Word?: unknown }).Word;
    const applyTypography = vi.spyOn(WordDocumentAdapter, 'applyParagraphTypography').mockResolvedValue(undefined);
    const doc = snapshot([paragraph({ index: 4, ooxmlAligned: false })]);
    vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(doc);
    (globalThis as { Word?: unknown }).Word = { run: vi.fn() };

    try {
      const result = await CorrectionService.apply([
        correctionFinding({ paragraphIndices: [4] })
      ]);
      expect(applyTypography).not.toHaveBeenCalled();
      expect(result.executionFailures).toBe(1);
      expect(result.attempts[0].message).toContain('could not be safely mapped');
    } finally {
      vi.restoreAllMocks();
      (globalThis as { Word?: unknown }).Word = previousWord;
    }
  });

  it('routes figure citation spelling through a prefix-only correction', async () => {
    const normalizeCitation = vi
      .spyOn(WordDocumentAdapter, 'normalizeFigureCitationSpelling')
      .mockResolvedValue({ applied: 1, skippedFieldBacked: [] });

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'FIGURE_CITATION_SPELLING_001',
          ruleName: 'Figure citation spelling',
          target: 'figureCaption',
          paragraphIndices: [3],
          property: 'crossref.figureCitationSpelling',
          rawExpected: true,
          expected: 'Full Figure/Figures wording',
          fixTextRanges: [{
            paragraphIndex: 3,
            searchText: 'Fig. 3',
            originalPrefix: 'Fig.',
            replacementPrefix: 'Figure',
            replacement: 'Figure 3',
            occurrenceIndex: 0
          }]
        })
      ]);

      expect(normalizeCitation).toHaveBeenCalledWith([{
        paragraphIndex: 3,
        searchText: 'Fig. 3',
        originalPrefix: 'Fig.',
        replacementPrefix: 'Figure',
        replacement: 'Figure 3',
        occurrenceIndex: 0
      }]);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('reports safe and field-backed figure citation corrections separately', async () => {
    const normalizeCitation = vi
      .spyOn(WordDocumentAdapter, 'normalizeFigureCitationSpelling')
      .mockResolvedValue({
        applied: 2,
        skippedFieldBacked: [{ paragraphIndex: 19, citation: 'Fig. 2' }]
      });

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'FIGURE_CITATION_SPELLING_001',
          ruleName: 'Figure citation spelling',
          target: 'figureCaption',
          paragraphIndices: [10, 19, 30],
          property: 'crossref.figureCitationSpelling',
          rawExpected: true,
          expected: 'Full Figure/Figures wording',
          fixTextRanges: []
        })
      ]);

      expect(normalizeCitation).toHaveBeenCalledWith([]);
      expect(result.figureCitationCorrection).toEqual({
        applied: 2,
        skippedFieldBacked: [{ paragraphIndex: 19, citation: 'Fig. 2' }]
      });
      expect(result.attempts[0].message).toContain('field-backed citation');
      expect(result.executedFixes).toBe(1);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('merges conclusion typography with an overlapping body-font finding', async () => {
    const applyTypography = vi
      .spyOn(WordDocumentAdapter, 'applyParagraphTypography')
      .mockResolvedValue(undefined);

    try {
      await CorrectionService.apply([
        correctionFinding({
          ruleId: 'BODY_FONT_001',
          paragraphIndices: [7],
          property: 'font.name',
          rawExpected: 'Times New Roman'
        }),
        correctionFinding({
          ruleId: 'CONCLUSION_BODY_FONT_001',
          ruleName: 'Conclusion body font',
          target: 'document',
          paragraphIndices: [7],
          property: 'document.conclusionTypography',
          rawExpected: true,
          expected: 'Times New Roman, 10 pt'
        })
      ]);

      expect(applyTypography).toHaveBeenCalledTimes(1);
      expect(applyTypography).toHaveBeenCalledWith(
        [7],
        { fontName: 'Times New Roman', fontSize: 10 }
      );
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('reports grouped typography execution failures instead of claiming success', async () => {
    vi.spyOn(WordDocumentAdapter, 'applyParagraphTypography')
      .mockRejectedValue(new Error('Office.js typography failure'));

    try {
      const result = await CorrectionService.apply([
        correctionFinding({ paragraphIndices: [2] })
      ]);

      expect(result.requestedFixes).toBe(1);
      expect(result.executedFixes).toBe(0);
      expect(result.executionFailures).toBe(1);
      expect(result.attempts[0].executionSucceeded).toBe(false);
      expect(result.attempts[0].message).toContain('Office.js typography failure');
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('continues and reports a partial correction batch after one operation fails', async () => {
    vi.spyOn(WordDocumentAdapter, 'applyParagraphTypography')
      .mockRejectedValue(new Error('Typography operation failed'));
    const applyEmailItalic = vi
      .spyOn(WordDocumentAdapter, 'applyAuthorTextItalic')
      .mockResolvedValue(undefined);

    try {
      const result = await CorrectionService.apply([
        correctionFinding({ paragraphIndices: [1] }),
        correctionFinding({
          ruleId: 'AUTHOR_EMAIL_ITALIC_001',
          ruleName: 'Author email italic',
          target: 'authorEmail',
          paragraphIndices: [2],
          property: 'font.italic',
          rawExpected: true,
          expected: 'Yes'
        })
      ]);

      expect(applyEmailItalic).toHaveBeenCalledWith([2], 'authorEmail', true);
      expect(result.requestedFixes).toBe(2);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(1);
      expect(result.attempts.find(attempt => attempt.ruleId === 'BODY_FONT_001')?.executionSucceeded).toBe(false);
      expect(result.attempts.find(attempt => attempt.ruleId === 'AUTHOR_EMAIL_ITALIC_001')?.executionSucceeded).toBe(true);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('keeps margin execution successful when immediate verification still fails', async () => {
    const applyMargin = vi.spyOn(WordDocumentAdapter, 'applySectionMargin').mockResolvedValue(undefined);
    vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(snapshot([], [{
      index: 0,
      marginBottomInches: 0.75
    }]));

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'MARGIN_BOTTOM_001',
          ruleName: 'Bottom margin',
          target: 'document',
          paragraphIndices: [],
          property: 'layout.marginBottomInches',
          rawExpected: 1,
          expected: '1.00 in'
        })
      ]);

      expect(result.attempts[0].executionSucceeded).toBe(true);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
      expect(result.attempts[0].message).toContain('verification still detected');
      expect(applyMargin).toHaveBeenCalledWith('bottom', 1);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('uses the expected value for both directions of author italic correction', async () => {
    const applyItalic = vi
      .spyOn(WordDocumentAdapter, 'applyAuthorTextItalic')
      .mockResolvedValue(undefined);

    try {
      const nonItalic = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'AUTHOR_EMAIL_ITALIC_001',
          ruleName: 'Author email italic',
          target: 'authorEmail',
          paragraphIndices: [4],
          property: 'font.italic',
          rawExpected: false,
          expected: 'No'
        })
      ]);
      const italic = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'AUTHOR_AFFILIATION_ITALIC_001',
          ruleName: 'Affiliation italic',
          target: 'authorAffiliation',
          paragraphIndices: [5],
          property: 'font.italic',
          rawExpected: true,
          expected: 'Yes'
        })
      ]);

      expect(applyItalic).toHaveBeenNthCalledWith(1, [4], 'authorEmail', false);
      expect(applyItalic).toHaveBeenNthCalledWith(2, [5], 'authorAffiliation', true);
      expect(nonItalic.executedFixes).toBe(1);
      expect(italic.executedFixes).toBe(1);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('uses a reviewed margin value rather than a hard-coded one', async () => {
    const applyMargin = vi.spyOn(WordDocumentAdapter, 'applySectionMargin').mockResolvedValue(undefined);
    vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(snapshot([], [{
      index: 0,
      marginTopInches: 0.75,
      marginRightInches: 0.75,
      marginBottomInches: 0.75,
      marginLeftInches: 0.75
    }]));

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'MARGIN_TOP_001',
          ruleName: 'Top margin',
          target: 'document',
          paragraphIndices: [],
          property: 'layout.marginTopInches',
          rawExpected: 0.75,
          expected: '0.75 in'
        }),
        correctionFinding({
          ruleId: 'MARGIN_BOTTOM_001',
          ruleName: 'Bottom margin',
          target: 'document',
          paragraphIndices: [],
          property: 'layout.marginBottomInches',
          rawExpected: 0.75,
          expected: '0.75 in'
        }),
        correctionFinding({
          ruleId: 'MARGIN_RIGHT_001',
          ruleName: 'Right margin',
          target: 'document',
          paragraphIndices: [],
          property: 'layout.marginRightInches',
          rawExpected: 0.75,
          expected: '0.75 in'
        }),
        correctionFinding({
          ruleId: 'MARGIN_LEFT_001',
          ruleName: 'Left margin',
          target: 'document',
          paragraphIndices: [],
          property: 'layout.marginLeftInches',
          rawExpected: 0.75,
          expected: '0.75 in'
        })
      ]);

      expect(applyMargin).toHaveBeenNthCalledWith(1, 'top', 0.75);
      expect(applyMargin).toHaveBeenNthCalledWith(2, 'bottom', 0.75);
      expect(applyMargin).toHaveBeenNthCalledWith(3, 'right', 0.75);
      expect(applyMargin).toHaveBeenNthCalledWith(4, 'left', 0.75);
      expect(result.executedFixes).toBe(4);
      expect(result.executionFailures).toBe(0);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('keeps line-spacing execution successful when the verification reread throws', async () => {
    vi.spyOn(WordDocumentAdapter, 'applySingleLineSpacing').mockResolvedValue(undefined);
    vi.spyOn(WordDocumentAdapter, 'snapshot').mockRejectedValue(new Error('OOXML reread failed'));

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'LINE_SPACING_001',
          ruleName: 'Single line spacing',
          property: 'paragraph.lineSpacing',
          paragraphIndices: [3],
          rawExpected: 12,
          expected: 'single'
        })
      ]);

      expect(result.attempts[0].executionSucceeded).toBe(true);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
      expect(result.attempts[0].message).toContain('post-correction verification failed');
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('routes table caption-to-table gap cleanup through the structural correction phase', async () => {
    const applyTableGap = vi
      .spyOn(WordDocumentAdapter, 'applyTableCaptionTableGaps')
      .mockResolvedValue(undefined);
    vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(snapshot([]));

    try {
      const result = await CorrectionService.apply([
        correctionFinding({
          ruleId: 'TABLE_CAPTION_TABLE_GAP_001',
          ruleName: 'Table caption directly precedes table',
          target: 'tableCaption',
          paragraphIndices: [24],
          property: 'structure.tableCaptionTableGap',
          rawExpected: true,
          expected: 'No blank paragraph'
        })
      ]);

      expect(applyTableGap).toHaveBeenCalledTimes(1);
      expect(result.requestedFixes).toBe(1);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('applies front-matter spacing as one verified before/after correction', async () => {
    const finding = validateRule('AUTHOR_NAME_SPACING_001', snapshot([
      paragraph({
        index: 6,
        text: 'Author Name',
        classification: 'authorName',
        spaceBefore: 0,
        spaceAfter: 0,
        ooxmlAligned: true
      })
    ]));
    const applySpacing = vi.spyOn(WordDocumentAdapter, 'applyParagraphSpacing').mockResolvedValue(undefined);
    const refreshed = snapshot([paragraph({
      index: 6,
      text: 'Author Name',
      classification: 'authorName',
      spaceBefore: 18,
      spaceAfter: 18,
      ooxmlAligned: true
    })]);
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot')
      .mockResolvedValueOnce(snapshot([paragraph({
        index: 6,
        text: 'Author Name',
        classification: 'authorName',
        spaceBefore: 0,
        spaceAfter: 0,
        ooxmlAligned: true
      })]))
      .mockResolvedValueOnce(refreshed);
    try {
      const result = await CorrectionService.apply([finding]);
      expect(applySpacing).toHaveBeenCalledWith([6], 18, 18);
      expect(result.attempts[0].executionSucceeded).toBe(true);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
    } finally {
      snapshotSpy.mockRestore();
      applySpacing.mockRestore();
    }
  });

  it('applies zero before/after spacing to affiliation and email paragraphs', async () => {
    const finding = validateRule('AUTHOR_AFFILIATION_SPACING_001', snapshot([
      paragraph({ index: 4, text: 'University of Example', classification: 'authorAffiliation', spaceBefore: 12, spaceAfter: 6, ooxmlAligned: true })
    ]));
    const applySpacing = vi.spyOn(WordDocumentAdapter, 'applyParagraphSpacing').mockResolvedValue(undefined);
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot')
      .mockResolvedValueOnce(snapshot([
        paragraph({ index: 4, text: 'University of Example', classification: 'authorAffiliation', spaceBefore: 12, spaceAfter: 6, ooxmlAligned: true })
      ]))
      .mockResolvedValueOnce(snapshot([
        paragraph({ index: 4, text: 'University of Example', classification: 'authorAffiliation', spaceBefore: 0, spaceAfter: 0, ooxmlAligned: true })
      ]));
    try {
      const result = await CorrectionService.apply([finding]);
      expect(applySpacing).toHaveBeenCalledWith([4], 0, 0);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
    } finally {
      snapshotSpy.mockRestore();
      applySpacing.mockRestore();
    }
  });

  it('routes affiliation-email separator correction through fresh structural verification', async () => {
    const affiliation = paragraph({ index: 0, text: 'Affiliation', classification: 'authorAffiliation' });
    const email = paragraph({ index: 1, text: 'email@example.com', classification: 'authorEmail' });
    const finding = validateRule('AFFILIATION_EMAIL_BLANK_001', snapshot([affiliation, email]));
    const applySeparator = vi.spyOn(WordDocumentAdapter, 'applyAffiliationEmailBlankLine').mockResolvedValue(undefined);
    const verified = paragraph({
      index: 1,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot')
      .mockResolvedValue(snapshot([affiliation, verified, paragraph({ index: 2, text: email.text, classification: 'authorEmail' })]));
    try {
      const result = await CorrectionService.apply([finding]);
      expect(applySeparator).toHaveBeenCalledTimes(1);
      expect(result.attempts[0].executionSucceeded).toBe(true);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
    } finally {
      snapshotSpy.mockRestore();
      applySeparator.mockRestore();
    }
  });

  it('routes front-matter extra blank cleanup through a fresh transition recheck', async () => {
    const blank = paragraph({
      index: 1,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const initial = snapshot([
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      blank,
      paragraph({ index: 2, text: 'Author One', classification: 'authorName' })
    ]);
    const final = snapshot([
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      paragraph({ index: 1, text: 'Author One', classification: 'authorName' })
    ]);
    const finding = validateRule('FRONTMATTER_EXTRA_BLANK_001', initial);
    const applyBlankCleanup = vi.spyOn(WordDocumentAdapter, 'applyFrontMatterExtraBlankLines').mockResolvedValue({
      removedBlankIndices: [1],
      alreadyResolvedTargets: 0
    });
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot')
      .mockResolvedValueOnce(initial)
      .mockResolvedValueOnce(final);
    try {
      const result = await CorrectionService.apply([finding]);
      expect(applyBlankCleanup).toHaveBeenCalledTimes(1);
      expect(applyBlankCleanup).toHaveBeenCalledWith(finding.frontMatterFixTargets);
      expect(result.attempts[0].executionSucceeded).toBe(true);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
    } finally {
      snapshotSpy.mockRestore();
      applyBlankCleanup.mockRestore();
    }
  });

  it('does not claim a stale front-matter blank correction executed', async () => {
    const finding = validateRule('FRONTMATTER_EXTRA_BLANK_001', snapshot([
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      paragraph({ index: 1, text: '', classification: 'other', isBlank: true, blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      } }),
      paragraph({ index: 2, text: 'Author One', classification: 'authorName' })
    ]));
    const applyBlankCleanup = vi.spyOn(WordDocumentAdapter, 'applyFrontMatterExtraBlankLines')
      .mockRejectedValue(new Error('No safely removable front-matter blank paragraph remains. Revalidate the document.'));
    try {
      const result = await CorrectionService.apply([finding]);
      expect(result.executedFixes).toBe(0);
      expect(result.executionFailures).toBe(1);
      expect(result.attempts[0].executionSucceeded).toBe(false);
      expect(result.attempts[0].message).toContain('Revalidate the document');
    } finally {
      applyBlankCleanup.mockRestore();
    }
  });

  it('uses a dedicated safe blank deletion operation for front-matter corrections', async () => {
    const source = await readFile('src/word/WordDocumentAdapter.ts', 'utf8');
    expect(source).toContain('static async removeSafeBlankParagraphs');
    expect(source).toContain('isStructurallySafeFrontMatterBlank(paragraph)');
    expect(source).toContain('static async applyFrontMatterExtraBlankLines');
  });

  it('does not apply an approved front-matter transition to a different newly introduced blank', async () => {
    const safeBlank = paragraph({
      index: 1,
      text: '',
      classification: 'other',
      isBlank: true,
      blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      }
    });
    const finding = validateRule('FRONTMATTER_EXTRA_BLANK_001', snapshot([
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      safeBlank,
      paragraph({ index: 2, text: 'Author One', classification: 'authorName' })
    ]));
    const changed = snapshot([
      paragraph({ index: 0, text: 'Paper Title', classification: 'title' }),
      paragraph({ index: 1, text: 'Author One', classification: 'authorName' }),
      paragraph({ index: 2, text: 'email@example.com', classification: 'authorEmail' }),
      paragraph({ index: 3, text: '', classification: 'other', isBlank: true, blankParagraph: {
        whitespaceOnly: true,
        hasManualBreak: false,
        hasExplicitPageBreak: false,
        hasLastRenderedPageBreak: false,
        hasSectionBreak: false,
        hasDrawing: false,
        hasField: false,
        hasEquation: false,
        hasNumbering: false,
        hasBookmarkContent: false,
        safeToDelete: true
      } }),
      paragraph({ index: 4, text: 'Abstract. Text', classification: 'abstract' })
    ]);
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(changed);
    const removeSpy = vi.spyOn(WordDocumentAdapter, 'removeSafeBlankParagraphs').mockResolvedValue(undefined);
    try {
      await expect(
        WordDocumentAdapter.applyFrontMatterExtraBlankLines(finding.frontMatterFixTargets)
      ).rejects.toThrow('no longer matches');
      expect(removeSpy).not.toHaveBeenCalled();
    } finally {
      snapshotSpy.mockRestore();
      removeSpy.mockRestore();
    }
  });

  it('applies Abstract prefix bold through a fresh exact-range correction', async () => {
    const finding = validateRule('ABSTRACT_PREFIX_BOLD_001', snapshot([
      paragraph({
        index: 8,
        text: 'Abstract. Body text',
        classification: 'abstract',
        abstractPrefixBold: false,
        abstractBodyBold: false,
        ooxmlAligned: true
      })
    ]));
    const applyPrefix = vi.spyOn(WordDocumentAdapter, 'applyAbstractPrefixBold').mockResolvedValue(undefined);
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot')
      .mockResolvedValueOnce(snapshot([paragraph({
        index: 8,
        text: 'Abstract. Body text',
        classification: 'abstract',
        abstractPrefixBold: false,
        abstractBodyBold: false,
        ooxmlAligned: true
      })]))
      .mockResolvedValueOnce(snapshot([paragraph({
        index: 8,
        text: 'Abstract. Body text',
        classification: 'abstract',
        abstractPrefixBold: true,
        abstractBodyBold: false,
        ooxmlAligned: true
      })]));
    try {
      const result = await CorrectionService.apply([finding]);
      expect(applyPrefix).toHaveBeenCalledWith([8]);
      expect(result.attempts[0].executionSucceeded).toBe(true);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
    } finally {
      snapshotSpy.mockRestore();
      applyPrefix.mockRestore();
    }
  });

  it('routes TABLE prefix bold through the exact-range correction path', async () => {
    const finding = validateRule('TABLE_PREFIX_BOLD_001', snapshot([
      paragraph({
        index: 3,
        text: 'TABLE 1. Results',
        classification: 'tableCaption',
        tablePrefixBold: false,
        ooxmlAligned: true
      })
    ]));
    const applyPrefix = vi.spyOn(WordDocumentAdapter, 'applyTableCaptionPrefixBold').mockResolvedValue(undefined);
    try {
      const result = await CorrectionService.apply([finding]);
      expect(applyPrefix).toHaveBeenCalledWith([3]);
      expect(result.attempts[0].executionSucceeded).toBe(true);
      expect(result.executedFixes).toBe(1);
      expect(result.executionFailures).toBe(0);
    } finally {
      applyPrefix.mockRestore();
    }
  });

  it('refuses an Abstract prefix bold fix when the validated text is stale', async () => {
    const finding = validateRule('ABSTRACT_PREFIX_BOLD_001', snapshot([
      paragraph({
        index: 8,
        text: 'Abstract. Original body',
        classification: 'abstract',
        abstractPrefixBold: false,
        ooxmlAligned: true
      })
    ]));
    const applyPrefix = vi.spyOn(WordDocumentAdapter, 'applyAbstractPrefixBold').mockResolvedValue(undefined);
    const snapshotSpy = vi.spyOn(WordDocumentAdapter, 'snapshot').mockResolvedValue(snapshot([paragraph({
      index: 8,
      text: 'Abstract. Changed body',
      classification: 'abstract',
      abstractPrefixBold: false,
      ooxmlAligned: true
    })]));
    try {
      const result = await CorrectionService.apply([finding]);
      expect(applyPrefix).not.toHaveBeenCalled();
      expect(result.executedFixes).toBe(0);
      expect(result.executionFailures).toBe(1);
      expect(result.attempts[0].message).toContain('text changed');
    } finally {
      snapshotSpy.mockRestore();
      applyPrefix.mockRestore();
    }
  });

  it('uses an exact Abstract prefix range rather than bolding the whole paragraph', async () => {
    const source = await readFile('src/word/WordDocumentAdapter.ts', 'utf8');
    expect(source).toContain('static async applyAbstractPrefixBold');
    expect(source).toContain('target.range.font.bold = true');
    expect(source).toContain('paragraphContentRange(entry.paragraph).search(prefix');
  });
});
