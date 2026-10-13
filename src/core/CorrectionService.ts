import { CorrectionAttempt, CorrectionResult, DocumentSnapshot, Finding, ParagraphSnapshot, SectionLayoutSnapshot, StrategyDefinition } from './models';
import { BUILT_IN_HEADING1_STYLE, extractCaptionPrefix, paragraphTextFingerprint, ParagraphTypographyOptions, WordDocumentAdapter } from '../word/WordDocumentAdapter';
import { analyzeAffiliationEmailBlankLine, analyzeFigureSpaceAbove, analyzeFrontMatterTransitions, analyzeTableCaptionTableGaps, findExtraBlankBodyParagraphIndices, findFigureCaptionBodyGapIssues, findFirstHeading1AfterAbstract, hasHeading1ParagraphStyle, isSafeAbstractPrefixBoldCandidate, isSafeHeading1StyleCandidate, ValidationEngine } from './ValidationEngine';

export function isVerifiedSingleLineSpacing(paragraph: ParagraphSnapshot): boolean {
  return paragraph.lineSpacingRule?.toLowerCase() === 'auto'
    && paragraph.lineSpacingMultiple !== null
    && Math.abs(paragraph.lineSpacingMultiple - 1) <= 0.05
    && paragraph.lineSpacingValue === 240;
}

export function verifySingleLineSpacing(paragraphs: ParagraphSnapshot[], indices: number[]): number[] {
  const byIndex = new Map(paragraphs.map(paragraph => [paragraph.index, paragraph]));
  return [...new Set(indices)].filter(index => {
    const paragraph = byIndex.get(index);
    return !paragraph || !isVerifiedSingleLineSpacing(paragraph);
  });
}

type SectionMarginSide = 'top' | 'right' | 'bottom' | 'left';

const SECTION_MARGIN_RULES: Readonly<Record<string, SectionMarginSide>> = {
  MARGIN_TOP_001: 'top',
  MARGIN_RIGHT_001: 'right',
  MARGIN_BOTTOM_001: 'bottom',
  MARGIN_LEFT_001: 'left'
};

const PAGE_SIZE_RULE_IDS = new Set(['PAGE_WIDTH_001', 'PAGE_HEIGHT_001']);

function isReferenceFormattingFinding(finding: Finding): boolean {
  return finding.autoFixable
    && finding.target === 'reference'
    && (finding.property === 'text.referenceNumbering' || finding.property.startsWith('reference.'));
}

function completeReferenceIndices(
  snapshot: DocumentSnapshot,
  requestedIndices: number[]
): number[] {
  const references = snapshot.paragraphs
    .filter(paragraph => paragraph.classification === 'reference')
    .map(paragraph => paragraph.index)
    .sort((a, b) => a - b);
  if (!references.length) {
    throw new Error('The References section could not be identified confidently. Reference formatting was withheld for manual review.');
  }

  const referenceSet = new Set(references);
  const stale = [...new Set(requestedIndices)].filter(index => !referenceSet.has(index));
  if (stale.length) {
    throw new Error('The validated reference paragraphs no longer match the current References section. Revalidate the document.');
  }

  const first = references[0];
  const last = references[references.length - 1];
  const ambiguous = snapshot.paragraphs.filter(paragraph =>
    paragraph.index >= first
    && paragraph.index <= last
    && paragraph.text.trim().length > 0
    && !referenceSet.has(paragraph.index)
  );
  if (ambiguous.length) {
    throw new Error('The References section contains an ambiguous paragraph boundary. Automatic reference formatting was withheld for manual review.');
  }
  return references;
}

function referenceContentSignature(text: string): string {
  return paragraphTextFingerprint(
    text.replace(/^\s*(?:\[\s*\d+\s*\]|\d+\s*[.)])\s+/, '')
  );
}

function referenceContentsUnchanged(before: DocumentSnapshot, after: DocumentSnapshot): boolean {
  const beforeReferences = before.paragraphs
    .filter(paragraph => paragraph.classification === 'reference')
    .sort((a, b) => a.index - b.index);
  const afterReferences = after.paragraphs
    .filter(paragraph => paragraph.classification === 'reference')
    .sort((a, b) => a.index - b.index);
  return beforeReferences.length === afterReferences.length
    && beforeReferences.every((paragraph, index) =>
      referenceContentSignature(paragraph.text) === referenceContentSignature(afterReferences[index].text)
    );
}

function sectionMarginValue(section: SectionLayoutSnapshot, side: SectionMarginSide): number | undefined {
  switch (side) {
    case 'top': return section.marginTopInches;
    case 'right': return section.marginRightInches;
    case 'bottom': return section.marginBottomInches;
    case 'left': return section.marginLeftInches;
  }
}

function sectionStructuralMetadataPreserved(
  beforeSections: SectionLayoutSnapshot[],
  afterSections: SectionLayoutSnapshot[]
): boolean {
  if (beforeSections.length !== afterSections.length) return false;
  return beforeSections.every(before => {
    const after = afterSections.find(candidate => candidate.index === before.index);
    if (!after) return false;

    if ((before.boundarySignature !== undefined || after.boundarySignature !== undefined)
      && before.boundarySignature !== after.boundarySignature) return false;
    if ((before.orientation !== undefined || after.orientation !== undefined)
      && before.orientation !== after.orientation) return false;

    return (['top', 'right', 'bottom', 'left'] as SectionMarginSide[]).every(side => {
      const beforeMargin = sectionMarginValue(before, side);
      const afterMargin = sectionMarginValue(after, side);
      if (beforeMargin === undefined && afterMargin === undefined) return true;
      return beforeMargin !== undefined
        && afterMargin !== undefined
        && Math.abs(beforeMargin - afterMargin) <= 0.01;
    });
  });
}

export class CorrectionService {
  static async apply(findings: Finding[], strategy?: StrategyDefinition): Promise<CorrectionResult> {
    const result: CorrectionResult = {
      verifications: [],
      attempts: [],
      requestedFixes: 0,
      executedFixes: 0,
      executionFailures: 0
    };

    // Structural deletion changes paragraph indices. Run it only after every
    // formatting/text correction has completed, so stored indices remain valid
    // for the rest of the selected fix batch.
    const structuralRuleIds = new Set([
      'EXTRA_BLANK_PARAGRAPH_001',
      'FIGURE_CAPTION_BODY_GAP_001',
      'FIGURE_LINE_SPACE_ABOVE_001',
      'TABLE_CAPTION_TABLE_GAP_001',
      'FRONTMATTER_EXTRA_BLANK_001',
      'AFFILIATION_EMAIL_BLANK_001'
    ]);
    const orderedFindings = [
      ...findings.filter(finding => !structuralRuleIds.has(finding.ruleId)),
      ...findings.filter(finding => finding.ruleId === 'EXTRA_BLANK_PARAGRAPH_001'),
      ...findings.filter(finding => finding.ruleId === 'FIGURE_CAPTION_BODY_GAP_001'),
      ...findings.filter(finding => finding.ruleId === 'FIGURE_LINE_SPACE_ABOVE_001'),
      ...findings.filter(finding => finding.ruleId === 'TABLE_CAPTION_TABLE_GAP_001'),
      ...findings.filter(finding => finding.ruleId === 'FRONTMATTER_EXTRA_BLANK_001'),
      ...findings.filter(finding => finding.ruleId === 'AFFILIATION_EMAIL_BLANK_001')
    ];

    // Merge overlapping generic typography fixes so a paragraph is formatted
    // once even when, for example, separate font and size findings are both
    // approved. Author-email/affiliation italic rules remain dedicated so
    // their association-marker exclusions are preserved.
    const dedicatedItalicRules = new Set([
      'AUTHOR_AFFILIATION_ITALIC_001',
      'AUTHOR_EMAIL_ITALIC_001'
    ]);
    const dedicatedAuthorTypographyRules = new Set([
      'AUTHOR_NAME_SIZE_001'
    ]);
    const typographyProperties = new Set(['font.name', 'font.size', 'font.bold', 'font.italic']);
    const requiresLayoutSafety = (finding: Finding): boolean =>
      finding.target === 'body'
      || finding.ruleId === 'LINE_SPACING_001'
      || finding.ruleId === 'TEXT_MULTIPLE_SPACES_001'
      || finding.ruleId === 'CONCLUSION_BODY_FONT_001';
    const typographyFindings = orderedFindings.filter(finding =>
      finding.autoFixable
      && (typographyProperties.has(finding.property) || finding.ruleId === 'CONCLUSION_BODY_FONT_001')
      && finding.rawExpected !== undefined
      && !dedicatedItalicRules.has(finding.ruleId)
      && !dedicatedAuthorTypographyRules.has(finding.ruleId)
      && !isReferenceFormattingFinding(finding)
    );
    const needsLayoutSafetySnapshot = orderedFindings.some(finding =>
      finding.autoFixable && requiresLayoutSafety(finding)
    );
    // Unit/integration callers may exercise the correction engine with mocked
    // adapter methods and no Office runtime. In Word, this guard always runs;
    // outside Word, leave the adapter mock in control of the test scenario.
    const officeRuntimeAvailable = typeof Word !== 'undefined' && typeof Word.run === 'function';
    const layoutSafetySnapshot: DocumentSnapshot | null = needsLayoutSafetySnapshot && officeRuntimeAvailable
      ? await WordDocumentAdapter.snapshot()
      : null;
    const layoutSafetyUnavailable = needsLayoutSafetySnapshot
      && officeRuntimeAvailable
      && (layoutSafetySnapshot === null
        || layoutSafetySnapshot.layout.ooxmlAvailable !== true
        || layoutSafetySnapshot.layout.paragraphMappingAligned !== true);
    const protectedLayoutIndices = new Set(
      (layoutSafetySnapshot?.paragraphs || [])
        .filter(paragraph => paragraph.layoutSensitive === true)
        .map(paragraph => paragraph.index)
    );
    const unreliableMappingIndices = new Set(
      (layoutSafetySnapshot?.paragraphs || [])
        .filter(paragraph => paragraph.ooxmlAligned === false)
        .map(paragraph => paragraph.index)
    );
    const safeGenericIndices = (indices: number[]): number[] =>
      layoutSafetyUnavailable
        ? []
        : [...new Set(indices)].filter(index =>
          !protectedLayoutIndices.has(index) && !unreliableMappingIndices.has(index)
        );
    const layoutSafetyMessage = layoutSafetyUnavailable
      ? 'Correction skipped because the document layout could not be inspected safely.'
      : 'Correction skipped because the affected paragraph belongs to a protected positioned/layout region.';
    const typographyIndicesByFinding = new Map<Finding, number[]>();
    for (const finding of typographyFindings) {
      const indices = requiresLayoutSafety(finding)
        ? safeGenericIndices(finding.paragraphIndices)
        : [...new Set(finding.paragraphIndices)];
      typographyIndicesByFinding.set(finding, indices);
    }
    const typographyByIndex = new Map<number, ParagraphTypographyOptions>();
    for (const finding of typographyFindings) {
      for (const index of typographyIndicesByFinding.get(finding) || []) {
        const options = typographyByIndex.get(index) || {};
        if (finding.ruleId === 'CONCLUSION_BODY_FONT_001') {
          options.fontName = 'Times New Roman';
          options.fontSize = 10;
        } else {
          if (finding.property === 'font.name') options.fontName = String(finding.rawExpected);
          if (finding.property === 'font.size') options.fontSize = Number(finding.rawExpected);
          if (finding.property === 'font.bold') options.bold = Boolean(finding.rawExpected);
          if (finding.property === 'font.italic') options.italic = Boolean(finding.rawExpected);
        }
        typographyByIndex.set(index, options);
      }
    }
    const groupedAttempts = typographyFindings.map(finding => {
      const attempt: CorrectionAttempt = {
        ruleId: finding.ruleId,
        paragraphIndices: [...new Set(finding.paragraphIndices)],
        attempted: true,
        executionSucceeded: false
      };
      const skipped = requiresLayoutSafety(finding)
        ? finding.paragraphIndices.filter(index => protectedLayoutIndices.has(index))
        : [];
      const mappingMismatchSkipped = requiresLayoutSafety(finding)
        ? finding.paragraphIndices.filter(index => unreliableMappingIndices.has(index))
        : [];
      if (skipped.length || mappingMismatchSkipped.length || (layoutSafetyUnavailable && requiresLayoutSafety(finding))) {
        const unsafeCount = new Set([...skipped, ...mappingMismatchSkipped]).size;
        attempt.message = layoutSafetyUnavailable
          ? layoutSafetyMessage
          : mappingMismatchSkipped.length
            ? `${unsafeCount} paragraph(s) could not be safely mapped to the current document layout and were skipped.`
            : `${unsafeCount} paragraph(s) belong to a protected positioned/layout region and were skipped.`;
      }
      result.attempts.push(attempt);
      result.requestedFixes += 1;
      return attempt;
    });
    if (typographyByIndex.size) {
      const typographyGroups = new Map<string, { indices: number[]; options: ParagraphTypographyOptions }>();
      const optionKeyByParagraph = new Map<number, string>();
      for (const [index, options] of typographyByIndex) {
        const key = [options.fontName || '', options.fontSize ?? '', options.bold ?? '', options.italic ?? ''].join('|');
        optionKeyByParagraph.set(index, key);
        const group = typographyGroups.get(key) || { indices: [], options };
        group.indices.push(index);
        typographyGroups.set(key, group);
      }

      const groupResults = new Map<string, { succeeded: boolean; message?: string }>();
      for (const [key, group] of typographyGroups) {
        try {
          await WordDocumentAdapter.applyParagraphTypography(group.indices, group.options);
          groupResults.set(key, { succeeded: true });
        } catch (error) {
          groupResults.set(key, {
            succeeded: false,
            message: error instanceof Error ? error.message : String(error)
          });
        }
      }

      for (const [finding, attempt] of groupedAttempts.map((attempt, index) => [typographyFindings[index], attempt] as const)) {
        const safeIndices = typographyIndicesByFinding.get(finding) || [];
        const keys = [...new Set(safeIndices.map(index => optionKeyByParagraph.get(index)))];
        if (!safeIndices.length) {
          result.executionFailures += 1;
          continue;
        }
        const failedKey = keys.find(key => !key || groupResults.get(key)?.succeeded !== true);
        if (!failedKey) {
          attempt.executionSucceeded = true;
          result.executedFixes += 1;
        } else {
          attempt.message = groupResults.get(failedKey)?.message
            || `Typography correction for ${finding.ruleId} did not complete.`;
          result.executionFailures += 1;
        }
      }
    } else {
      for (const finding of typographyFindings) {
        if (!(typographyIndicesByFinding.get(finding) || []).length) {
          result.executionFailures += 1;
        }
      }
    }
    const groupedTypographyFindings = new Set(typographyFindings);

    // Width and height belong to one section-level paper-size correction.
    // Keep one attempt per selected finding for the task pane, but execute
    // one host operation so orientation and pagination are handled together.
    const pageSizeFindings = orderedFindings.filter(finding => PAGE_SIZE_RULE_IDS.has(finding.ruleId));
    const pageSizeAttempts = new Map<Finding, CorrectionAttempt>();
    for (const finding of pageSizeFindings) {
      const attempt: CorrectionAttempt = {
        ruleId: finding.ruleId,
        paragraphIndices: [],
        attempted: true,
        executionSucceeded: false
      };
      pageSizeAttempts.set(finding, attempt);
      result.attempts.push(attempt);
      result.requestedFixes += 1;
    }
    if (pageSizeFindings.length) {
      try {
        const before = officeRuntimeAvailable
          ? await WordDocumentAdapter.snapshot()
          : null;
        const beforeSections = before?.layout.sections || [];
        const beforeSectionIndices = new Set(beforeSections.map(section => section.index));
        await WordDocumentAdapter.applyLetterPageSize();
        const reread = officeRuntimeAvailable ? await WordDocumentAdapter.snapshot() : null;
        const rereadSections = reread?.layout.sections || [];
        const sectionMetadataIncomplete = !before
          || beforeSections.length === 0
          || !reread
          || rereadSections.length !== beforeSections.length
          || [...beforeSectionIndices].some(index => !rereadSections.some(section => section.index === index))
          || !sectionStructuralMetadataPreserved(beforeSections, rereadSections);
        const invalid = sectionMetadataIncomplete
          ? [null]
          : rereadSections.filter(section =>
            section.pageWidthInches === undefined
            || section.pageHeightInches === undefined
            || (() => {
              const landscape = section.orientation === 'landscape'
                || section.pageWidthInches! > section.pageHeightInches!;
              const expectedWidth = landscape ? 11 : 8.5;
              const expectedHeight = landscape ? 8.5 : 11;
              return Math.abs(section.pageWidthInches! - expectedWidth) > 0.02
                || Math.abs(section.pageHeightInches! - expectedHeight) > 0.02;
            })()
          );
        for (const finding of pageSizeFindings) {
          const attempt = pageSizeAttempts.get(finding)!;
          if (invalid.length) {
            attempt.message = sectionMetadataIncomplete
              ? 'US Letter correction executed, but section metadata was unavailable for verification. Paper size changes can alter pagination, table wrapping, and figure placement.'
              : `US Letter correction executed, but verification still detected ${invalid.length} section(s) with non-Letter dimensions. Paper size changes can alter pagination, table wrapping, and figure placement.`;
            result.executionFailures += 1;
          } else {
            attempt.executionSucceeded = true;
            result.executedFixes += 1;
          }
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        for (const finding of pageSizeFindings) {
          pageSizeAttempts.get(finding)!.message = message;
          result.executionFailures += 1;
        }
      }
    }

    // Reference numbering, style, typography, and list positions are one
    // interdependent correction. The individual findings remain visible and
    // countable, while the adapter applies the union in a single transaction.
    const referenceFindings = orderedFindings.filter(isReferenceFormattingFinding);
    const referenceAttempts = new Map<Finding, CorrectionAttempt>();
    for (const finding of referenceFindings) {
      const attempt: CorrectionAttempt = {
        ruleId: finding.ruleId,
        paragraphIndices: [...new Set(finding.paragraphIndices)],
        attempted: true,
        executionSucceeded: false
      };
      referenceAttempts.set(finding, attempt);
      result.attempts.push(attempt);
      result.requestedFixes += 1;
    }
    if (referenceFindings.length) {
      const requestedReferenceIndices = [...new Set(referenceFindings.flatMap(finding =>
        finding.fixParagraphIndices ?? finding.paragraphIndices
      ))].sort((a, b) => a - b);
      const numberingRequested = referenceFindings.some(finding =>
        finding.property === 'text.referenceNumbering'
        || finding.property.startsWith('reference.list.')
      );
      const referenceProperties = [...new Set(referenceFindings.map(finding => finding.property))];
      try {
        const freshBeforeCorrection = officeRuntimeAvailable
          ? await WordDocumentAdapter.snapshot()
          : null;
        const referenceIndices = freshBeforeCorrection
          ? completeReferenceIndices(freshBeforeCorrection, requestedReferenceIndices)
          : requestedReferenceIndices;
        for (const finding of referenceFindings) {
          referenceAttempts.get(finding)!.paragraphIndices = referenceIndices;
        }
        await WordDocumentAdapter.applyReferenceFormatting(referenceIndices, {
          applyNumbering: numberingRequested,
          properties: referenceProperties
        });

        const freshAfterCorrection = officeRuntimeAvailable
          ? await WordDocumentAdapter.snapshot()
          : null;
        const validationByRule = strategy && freshAfterCorrection
          ? new Map(ValidationEngine.validate(strategy, freshAfterCorrection).findings.map(finding => [finding.ruleId, finding]))
          : null;
        const verificationEvidenceAvailable = !strategy
          || Boolean(freshBeforeCorrection && freshAfterCorrection && validationByRule);
        const referenceContentValid = freshBeforeCorrection && freshAfterCorrection
          ? referenceContentsUnchanged(freshBeforeCorrection, freshAfterCorrection)
          : true;
        for (const finding of referenceFindings) {
          const after = validationByRule?.get(finding.ruleId);
          const requiredVerificationRuleIds = new Set([finding.ruleId]);
          if (finding.property === 'reference.style.leftIndent'
            || finding.property === 'reference.style.hangingIndent') {
            // These two properties are one coupled style layer.  A fix for
            // either one must prove that both rules pass after Word sync.
            for (const ruleId of ['REFERENCES_LEFT_INDENT_001', 'REFERENCES_HANGING_INDENT_001']) {
              if (strategy?.rules.some(rule => rule.id === ruleId)) requiredVerificationRuleIds.add(ruleId);
            }
          }
          const unverifiedRuleIds = [...requiredVerificationRuleIds].filter(ruleId =>
            validationByRule?.get(ruleId)?.status !== 'passed'
          );
          const verified = !strategy
            || (verificationEvidenceAvailable && after?.status === 'passed'
              && unverifiedRuleIds.length === 0
              && referenceContentValid);
          const attempt = referenceAttempts.get(finding)!;
          attempt.executionSucceeded = true;
          attempt.verificationSucceeded = strategy ? verified : undefined;
          result.executedFixes += 1;
          if (!verified) {
            attempt.message = !referenceContentValid
              ? 'Reference formatting executed, but reference count or text changed during correction. Manual review is required.'
              : !verificationEvidenceAvailable
                ? 'Reference formatting executed, but post-correction validation evidence was unavailable. Manual review is required.'
              : `Reference formatting executed, but ${unverifiedRuleIds.join(', ') || finding.ruleId} was not explicitly passed after revalidation.`;
            result.executionFailures += 1;
          }
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        for (const finding of referenceFindings) {
          referenceAttempts.get(finding)!.verificationSucceeded = false;
          referenceAttempts.get(finding)!.message = message;
          result.executionFailures += 1;
        }
      }
    }

    for (const finding of orderedFindings) {
      if (!finding.autoFixable) continue;
      if (groupedTypographyFindings.has(finding)) continue;
      if (PAGE_SIZE_RULE_IDS.has(finding.ruleId) || referenceAttempts.has(finding)) continue;

      const attempt: CorrectionAttempt = {
        ruleId: finding.ruleId,
        paragraphIndices: [...new Set(finding.paragraphIndices)],
        attempted: true,
        executionSucceeded: false
      };
      result.attempts.push(attempt);
      result.requestedFixes += 1;

      const markExecuted = (): void => {
        if (!attempt.executionSucceeded) {
          attempt.executionSucceeded = true;
          result.executedFixes += 1;
        }
      };

      try {
        if (finding.ruleId === 'LINE_SPACING_001') {
        const requestedParagraphIndices = [...new Set(finding.paragraphIndices)];
        const safeParagraphIndices = safeGenericIndices(requestedParagraphIndices);
        const skippedCount = requestedParagraphIndices.length - safeParagraphIndices.length;
        if (!safeParagraphIndices.length) {
          attempt.message = layoutSafetyMessage;
          result.executionFailures += 1;
          continue;
        }
        await WordDocumentAdapter.applySingleLineSpacing(safeParagraphIndices);
          markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        const failedParagraphIndices = verifySingleLineSpacing(
          reread.paragraphs,
          safeParagraphIndices
        );
        const failed = new Set(failedParagraphIndices);
        const verifiedParagraphIndices = safeParagraphIndices.filter(index => !failed.has(index));
        result.verifications.push({
          ruleId: finding.ruleId,
          requestedParagraphIndices: safeParagraphIndices,
          verifiedParagraphIndices,
          failedParagraphIndices
        });
        const messages: string[] = [];
        if (skippedCount > 0) messages.push(`${skippedCount} paragraph(s) skipped because their layout mapping was not safe.`);
        if (failedParagraphIndices.length) messages.push(`${failedParagraphIndices.length} paragraph fix(es) failed verification.`);
        if (messages.length) attempt.message = messages.join(' ');
        continue;
        }

      const marginSide = SECTION_MARGIN_RULES[finding.ruleId];
      if (marginSide) {
        const expectedMargin = Number(finding.rawExpected);
        if (!Number.isFinite(expectedMargin) || expectedMargin < 0) {
          throw new Error(`${finding.ruleId} requires a non-negative numeric expected margin.`);
        }
        await WordDocumentAdapter.applySectionMargin(marginSide, expectedMargin);
        markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        const invalidSections = (reread.layout.sections || []).filter(section =>
          sectionMarginValue(section, marginSide) === undefined
          || sectionMarginValue(section, marginSide) === null
          || Math.abs((sectionMarginValue(section, marginSide) as number) - expectedMargin) > 0.02
        );
        if (invalidSections.length) {
          attempt.message = `${marginSide[0].toUpperCase()}${marginSide.slice(1)}-margin correction executed, but verification still detected ${invalidSections.length} invalid section(s).`;
        }
        continue;
      }

      if (finding.ruleId === 'PAGE_NUMBER_FORBIDDEN_001') {
        await WordDocumentAdapter.removePageNumberFields();
        markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        if (reread.layout.containsPageNumberField === true) {
          attempt.message = 'Page-number field removal executed, but verification still found a PAGE, NUMPAGES, or SECTIONPAGES field.';
        }
        continue;
      }

      if (finding.ruleId === 'EXTRA_BLANK_PARAGRAPH_001') {
        const requestedParagraphIndices = [...new Set(finding.paragraphIndices)];
        await WordDocumentAdapter.removeBlankParagraphs(requestedParagraphIndices);
        markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        const remaining = findExtraBlankBodyParagraphIndices(reread.paragraphs);
        if (remaining.length) {
          attempt.message = `Blank-paragraph correction executed, but verification still found ${remaining.length} safely removable blank paragraph(s).`;
        }
        continue;
      }

      if (finding.ruleId === 'FIGURE_CAPTION_BODY_GAP_001') {
        await WordDocumentAdapter.applyFigureCaptionBodyGaps();
        markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        const remaining = findFigureCaptionBodyGapIssues(reread.paragraphs)
          .filter(issue => !issue.blockedByStructuralParagraph);
        if (remaining.length) {
          attempt.message = `Figure-caption gap correction executed, but verification still found ${remaining.length} applicable caption gap(s).`;
        }
        continue;
      }

      if (finding.ruleId === 'FIGURE_LINE_SPACE_ABOVE_001') {
        await WordDocumentAdapter.applyFigureSpaceAbove();
        markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        const remaining = analyzeFigureSpaceAbove(reread)
          .issues
          .filter(issue => !issue.blockedByStructuralParagraph);
        if (remaining.length) {
          attempt.message = `Figure-spacing correction executed, but verification still found ${remaining.length} applicable figure gap(s).`;
        }
        continue;
      }

      if (finding.ruleId === 'TABLE_CAPTION_TABLE_GAP_001') {
        await WordDocumentAdapter.applyTableCaptionTableGaps();
        markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        const remaining = analyzeTableCaptionTableGaps(reread)
          .issues
          .filter(issue => !issue.blockedByStructuralParagraph);
        if (remaining.length) {
          attempt.message = `Table-caption gap correction executed, but verification still found ${remaining.length} applicable caption-to-table gap(s).`;
        }
        continue;
      }

      if (finding.ruleId === 'AFFILIATION_EMAIL_BLANK_001') {
        await WordDocumentAdapter.applyAffiliationEmailBlankLine();
        markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        const remaining = analyzeAffiliationEmailBlankLine(reread.paragraphs);
        attempt.verificationSucceeded = remaining.state === 'passed';
        if (remaining.state !== 'passed') {
          attempt.message = `Affiliation-email separator correction executed, but verification found state '${remaining.state}'.`;
          result.executionFailures += 1;
        }
        continue;
      }

      if (finding.ruleId === 'FRONTMATTER_EXTRA_BLANK_001') {
        const outcome = await WordDocumentAdapter.applyFrontMatterExtraBlankLines(finding.frontMatterFixTargets || []);
        markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        const remainingApprovedTargets = (finding.frontMatterFixTargets || []).filter(target =>
          !analyzeFrontMatterTransitions(reread.paragraphs).some(transition =>
            transition.fromType === target.fromType
            && transition.toType === target.toType
            && transition.fromTextFingerprint === target.fromTextFingerprint
            && transition.toTextFingerprint === target.toTextFingerprint
            && transition.expectedBlankCount === 0
            && transition.interveningNonBlankIndices.length === 0
            && transition.blankIndices.length === 0
          )
        );
        if (outcome.alreadyResolvedTargets || remainingApprovedTargets.length) {
          const messages: string[] = [];
          if (outcome.alreadyResolvedTargets) messages.push(`${outcome.alreadyResolvedTargets} transition(s) were already resolved`);
          if (remainingApprovedTargets.length) messages.push(`${remainingApprovedTargets.length} approved transition(s) failed post-correction verification`);
          attempt.message = messages.join('; ');
        }
        continue;
      }

      if (finding.ruleId === 'AUTHOR_NAME_SPACING_001'
        || finding.ruleId === 'AUTHOR_AFFILIATION_SPACING_001'
        || finding.ruleId === 'AUTHOR_EMAIL_SPACING_001'
        || finding.ruleId === 'ABSTRACT_SPACING_001') {
        const current = await WordDocumentAdapter.snapshot();
        const expectedTexts = finding.fixParagraphTexts || {};
        const requestedIndices = [...new Set(finding.paragraphIndices)];
        const safeIndices = requestedIndices.filter(index => {
          const paragraph = current.paragraphs.find(candidate => candidate.index === index);
          if (!paragraph
            || paragraph.classification !== finding.target
            || paragraph.tableNestingLevel !== 0
            || paragraph.layoutSensitive === true
            || paragraph.ooxmlAligned === false
            || current.layout.paragraphMappingAligned === false) {
            return false;
          }
          const expectedText = expectedTexts[index];
          return expectedText === undefined
            || paragraphTextFingerprint(paragraph.text) === paragraphTextFingerprint(expectedText);
        });
        const skippedCount = requestedIndices.length - safeIndices.length;
        if (!safeIndices.length) {
          attempt.message = 'Front-matter spacing correction skipped because the paragraph text or layout mapping changed. Revalidate the document.';
          result.executionFailures += 1;
          continue;
        }
        const spacing = Number(finding.rawExpected);
        await WordDocumentAdapter.applyParagraphSpacing(safeIndices, spacing, spacing);
        markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        const rereadByIndex = new Map(reread.paragraphs.map(paragraph => [paragraph.index, paragraph]));
        const unresolved = safeIndices.filter(index => {
          const paragraph = rereadByIndex.get(index);
          return !paragraph
            || paragraph.spaceBefore === null
            || paragraph.spaceBefore === undefined
            || paragraph.spaceAfter === null
            || paragraph.spaceAfter === undefined
            || Math.abs(paragraph.spaceBefore - spacing) > 0.75
            || Math.abs(paragraph.spaceAfter - spacing) > 0.75;
        });
        const messages: string[] = [];
        if (skippedCount) messages.push(`${skippedCount} paragraph(s) skipped because their text or layout mapping was no longer safe.`);
        if (unresolved.length) messages.push(`${unresolved.length} front-matter spacing correction(s) failed verification.`);
        if (messages.length) attempt.message = messages.join(' ');
        continue;
      }

      if (finding.ruleId === 'FIRST_HEADING1_SPACE_BEFORE_001') {
        const current = await WordDocumentAdapter.snapshot();
        const expectedIndex = finding.paragraphIndices[0];
        const expectedText = finding.fixParagraphTexts?.[expectedIndex];
        const firstHeading = findFirstHeading1AfterAbstract(current.paragraphs);
        const safe = firstHeading !== null
          && firstHeading.index === expectedIndex
          && firstHeading.classification === 'heading1'
          && firstHeading.tableNestingLevel === 0
          && firstHeading.layoutSensitive !== true
          && firstHeading.ooxmlAligned !== false
          && current.layout.ooxmlAvailable !== false
          && current.layout.paragraphMappingAligned !== false
          && (expectedText === undefined
            || paragraphTextFingerprint(firstHeading.text) === paragraphTextFingerprint(expectedText));
        if (!safe || !firstHeading) {
          attempt.message = 'First Heading 1 spacing correction skipped because the Abstract/heading structure or text changed. Revalidate the document.';
          result.executionFailures += 1;
          continue;
        }
        await WordDocumentAdapter.applyParagraphSpacingBefore([firstHeading.index], 0);
        markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        const verifiedHeading = findFirstHeading1AfterAbstract(reread.paragraphs);
        const verified = verifiedHeading !== null
          && verifiedHeading.index === expectedIndex
          && paragraphTextFingerprint(verifiedHeading.text) === paragraphTextFingerprint(firstHeading.text)
          && verifiedHeading.spaceBefore !== null
          && verifiedHeading.spaceBefore !== undefined
          && Math.abs(verifiedHeading.spaceBefore) <= 0.75;
        if (!verified) {
          attempt.message = 'First Heading 1 spacing correction executed, but verification did not find 0 pt space before on the same heading.';
        }
        continue;
      }

      if (finding.ruleId === 'ABSTRACT_PREFIX_BOLD_001') {
        const current = await WordDocumentAdapter.snapshot();
        const expectedTexts = finding.fixParagraphTexts || {};
        const requestedIndices = [...new Set(finding.fixParagraphIndices ?? finding.paragraphIndices)];
        const safeIndices = requestedIndices.filter(index => {
          const paragraph = current.paragraphs.find(candidate => candidate.index === index);
          if (!paragraph
            || !isSafeAbstractPrefixBoldCandidate(
              paragraph,
              current.layout.paragraphMappingAligned !== false
            )) {
            return false;
          }
          const expectedText = expectedTexts[index];
          return expectedText === undefined
            || paragraphTextFingerprint(paragraph.text) === paragraphTextFingerprint(expectedText);
        });
        const skippedCount = requestedIndices.length - safeIndices.length;
        if (!safeIndices.length) {
          attempt.message = 'Abstract prefix bold correction skipped because the exact prefix, field mapping, or paragraph text changed. Revalidate the document.';
          result.executionFailures += 1;
          continue;
        }
        await WordDocumentAdapter.applyAbstractPrefixBold(safeIndices);
        markExecuted();
        const reread = await WordDocumentAdapter.snapshot();
        const rereadByIndex = new Map(reread.paragraphs.map(paragraph => [paragraph.index, paragraph]));
        const unresolved = safeIndices.filter(index => rereadByIndex.get(index)?.abstractPrefixBold !== true);
        const messages: string[] = [];
        if (skippedCount) messages.push(`${skippedCount} Abstract prefix correction(s) skipped because the prefix or paragraph mapping was no longer safe.`);
        if (unresolved.length) messages.push(`${unresolved.length} Abstract prefix bold correction(s) failed verification.`);
        if (messages.length) attempt.message = messages.join(' ');
        continue;
      }

      if (finding.ruleId === 'FIGURE_CITATION_SPELLING_001') {
        const citationResult = await WordDocumentAdapter.normalizeFigureCitationSpelling(finding.fixTextRanges || []);
        result.figureCitationCorrection = citationResult;
        markExecuted();
        if (citationResult.skippedFieldBacked.length) {
          attempt.message = `${citationResult.applied} figure citation(s) corrected; ${citationResult.skippedFieldBacked.length} field-backed citation(s) require manual review.`;
        }
        continue;
      }

      if (finding.ruleId === 'TEXT_MULTIPLE_SPACES_001') {
        const requestedParagraphIndices = [...new Set(finding.paragraphIndices)];
        const safeParagraphIndices = safeGenericIndices(finding.paragraphIndices);
        const skippedCount = requestedParagraphIndices.length - safeParagraphIndices.length;
        if (!safeParagraphIndices.length) {
          attempt.message = layoutSafetyMessage;
          result.executionFailures += 1;
          continue;
        }
        await WordDocumentAdapter.normalizeParagraphWhitespace(safeParagraphIndices, 'internal');
        markExecuted();
        if (skippedCount > 0) {
          attempt.message = `${skippedCount} paragraph(s) skipped because their layout mapping was not safe.`;
        }
        continue;
      }

      if (finding.ruleId === 'EQUATION_TABS_001') {
        await WordDocumentAdapter.applyEquationTabs(finding.paragraphIndices);
        markExecuted();
        continue;
      }

      if (finding.ruleId === 'TABLE_TOP_RULE_001' || finding.ruleId === 'TABLE_BOTTOM_RULE_001' || finding.ruleId === 'TABLE_HEADER_RULE_001'
        || finding.ruleId === 'TABLE_NO_VERTICAL_RULES_001' || finding.ruleId === 'TABLE_NO_EXTRA_HORIZONTAL_RULES_001') {
        await WordDocumentAdapter.applyTableRule(
          finding.paragraphIndices,
          finding.ruleId === 'TABLE_TOP_RULE_001' ? 'top'
            : finding.ruleId === 'TABLE_BOTTOM_RULE_001' ? 'bottom'
                : finding.ruleId === 'TABLE_HEADER_RULE_001' ? 'header'
                  : finding.ruleId === 'TABLE_NO_VERTICAL_RULES_001' ? 'vertical' : 'extraHorizontal',
          finding.fixTableIndices,
          finding.fixTableHeaderRowCounts,
          finding.fixTableCaptionRowIndices,
          finding.fixTableHeaderStartRowIndices,
          finding.fixTableHeaderEndRowIndices
        );
        markExecuted();
        continue;
      }

      if (finding.ruleId === 'FIGURE_CAPTION_LEADING_WHITESPACE_001'
        || finding.ruleId === 'TABLE_CAPTION_LEADING_WHITESPACE_001') {
        await WordDocumentAdapter.normalizeParagraphWhitespace(finding.paragraphIndices, 'leading');
        markExecuted();
        continue;
      }

      // Dedicated structural/text corrections. These are intentionally kept
      // out of the generic paragraph-property path because they alter content
      // or non-body document parts.
      switch (finding.ruleId) {
        case 'HEADING1_STYLE_001': {
          const current = await WordDocumentAdapter.snapshot();
          const byIndex = new Map(current.paragraphs.map(paragraph => [paragraph.index, paragraph]));
          const expectedTexts = finding.fixParagraphTexts || {};
          const requestedIndices = [...new Set(finding.paragraphIndices)];
          const safeIndices = requestedIndices.filter(index => {
            const paragraph = byIndex.get(index);
            if (!paragraph || !isSafeHeading1StyleCandidate(paragraph) || hasHeading1ParagraphStyle(paragraph)) return false;
            if (current.layout.ooxmlAvailable !== true || current.layout.paragraphMappingAligned !== true) return false;
            const expectedText = expectedTexts[index];
            return expectedText === undefined
              || paragraphTextFingerprint(paragraph.text) === paragraphTextFingerprint(expectedText);
          });
          const skippedCount = requestedIndices.length - safeIndices.length;
          if (!safeIndices.length) {
            attempt.message = 'Heading 1 style correction skipped because the paragraph text or layout mapping changed. Revalidate the document.';
            result.executionFailures += 1;
            continue;
          }
          await WordDocumentAdapter.applyBuiltInParagraphStyle(safeIndices, BUILT_IN_HEADING1_STYLE);
          markExecuted();
          const reread = await WordDocumentAdapter.snapshot();
          const rereadByIndex = new Map(reread.paragraphs.map(paragraph => [paragraph.index, paragraph]));
          const unresolved = safeIndices.filter(index => !hasHeading1ParagraphStyle(rereadByIndex.get(index) || {} as ParagraphSnapshot));
          const messages: string[] = [];
          if (skippedCount) messages.push(`${skippedCount} heading paragraph(s) skipped because their text or layout mapping was no longer safe.`);
          if (unresolved.length) messages.push(`${unresolved.length} Heading 1 style correction(s) failed verification.`);
          if (messages.length) attempt.message = messages.join(' ');
          continue;
        }
        case 'AUTHOR_NAME_SIZE_001':
          await WordDocumentAdapter.applyAuthorNameTypographyPreservingMarkers(
            finding.paragraphIndices,
            { fontSize: Number(finding.rawExpected) }
          );
          markExecuted();
          continue;
        case 'AUTHOR_AFFILIATION_SUPERSCRIPT_001':
          await WordDocumentAdapter.applyAuthorAssociationSuperscripts(finding.paragraphIndices, 'authorAffiliation');
          markExecuted();
          continue;
        case 'AUTHOR_EMAIL_SUPERSCRIPT_001':
          await WordDocumentAdapter.applyAuthorAssociationSuperscripts(finding.paragraphIndices, 'authorEmail');
          markExecuted();
          continue;
        case 'HEADER_FORBIDDEN_001':
          await WordDocumentAdapter.clearHeaders();
          markExecuted();
          continue;
        case 'FOOTER_FORBIDDEN_001':
          await WordDocumentAdapter.clearFooters();
          markExecuted();
          continue;
        case 'TITLE_CASE_001':
          await WordDocumentAdapter.applyTitleCase(finding.paragraphIndices);
          markExecuted();
          continue;
        case 'FIGURE_CAPTION_PREFIX_001':
          await WordDocumentAdapter.normalizeFigureCaptionPrefixes(
            finding.paragraphIndices,
            finding.fixParagraphIndices ?? finding.paragraphIndices
          );
          markExecuted();
          continue;
        case 'TABLE_CAPTION_PREFIX_001':
          await WordDocumentAdapter.normalizeTableCaptionPrefixes(
            finding.fixParagraphIndices ?? finding.paragraphIndices
          );
          markExecuted();
          continue;
        case 'TABLE_PREFIX_BOLD_001':
          await WordDocumentAdapter.applyTableCaptionPrefixBold(
            finding.fixParagraphIndices ?? finding.paragraphIndices
          );
          markExecuted();
          continue;
        case 'REFERENCES_NUMBERING_001':
          await WordDocumentAdapter.normalizeReferenceNumbering(
            finding.fixParagraphIndices ?? finding.paragraphIndices
          );
          markExecuted();
          continue;
        case 'AUTHOR_AFFILIATION_ITALIC_001':
          if (typeof finding.rawExpected !== 'boolean') {
            throw new Error('AUTHOR_AFFILIATION_ITALIC_001 requires a boolean expected value.');
          }
          await WordDocumentAdapter.applyAuthorTextItalic(
            finding.paragraphIndices,
            'authorAffiliation',
            finding.rawExpected
          );
          markExecuted();
          continue;
        case 'AUTHOR_EMAIL_ITALIC_001':
          if (typeof finding.rawExpected !== 'boolean') {
            throw new Error('AUTHOR_EMAIL_ITALIC_001 requires a boolean expected value.');
          }
          await WordDocumentAdapter.applyAuthorTextItalic(
            finding.paragraphIndices,
            'authorEmail',
            finding.rawExpected
          );
          markExecuted();
          continue;
      }

      if (finding.rawExpected === undefined) {
        throw new Error(`No deterministic correction is mapped for ${finding.ruleId}.`);
      }

      if (finding.property === 'text.allCaps') {
        const safeParagraphIndices = requiresLayoutSafety(finding)
          ? safeGenericIndices(finding.paragraphIndices)
          : [...new Set(finding.paragraphIndices)];
        if (!safeParagraphIndices.length) {
          attempt.message = layoutSafetyMessage;
          result.executionFailures += 1;
          continue;
        }
        await WordDocumentAdapter.applyParagraphProperty(
          safeParagraphIndices,
          'text.uppercase',
          true
        );
        markExecuted();
        continue;
      }

      const genericParagraphIndices = requiresLayoutSafety(finding)
        ? safeGenericIndices(finding.paragraphIndices)
        : [...new Set(finding.paragraphIndices)];
      if (!genericParagraphIndices.length) {
        attempt.message = layoutSafetyMessage;
        result.executionFailures += 1;
        continue;
      }
      await WordDocumentAdapter.applyParagraphProperty(
        genericParagraphIndices,
        finding.property,
        finding.rawExpected
      );
      markExecuted();
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        attempt.message = attempt.executionSucceeded
          ? `Correction executed, but post-correction verification failed: ${message}`
          : message;
        if (!attempt.executionSucceeded) result.executionFailures += 1;
      }
    }

    return result;
  }
}
