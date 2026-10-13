import type {
  CanonicalSection,
  SectionRecognitionConfidence
} from './SectionOntology';

export type FindingStatus = 'passed' | 'failed' | 'manual_review' | 'checklist' | 'not_applicable' | 'not_evaluated';
export type Severity = 'critical' | 'major' | 'minor' | 'info';
export type ReviewKind = 'source_conflict' | 'manual' | 'submission' | 'technical';
export type RuleTarget =
  | 'document'
  | 'title'
  | 'authorName'
  | 'authorAffiliation'
  | 'authorEmail'
  | 'abstract'
  | 'body'
  | 'heading1'
  | 'heading2'
  | 'heading3'
  | 'figureCaption'
  | 'tableCaption'
  | 'equation'
  | 'referencesHeading'
  | 'reference';

export type FrontMatterType =
  | 'title'
  | 'authorName'
  | 'authorAffiliation'
  | 'authorEmail'
  | 'abstract';

export interface FrontMatterTransitionFixTarget {
  fromType: FrontMatterType;
  toType: FrontMatterType;
  fromIndex: number;
  toIndex: number;
  fromTextFingerprint: string;
  toTextFingerprint: string;
  blankIndices: number[];
  expectedBlankCount: number;
}

export interface ConflictAlternative {
  source: string;
  value: string | number | boolean;
  note?: string;
}

export interface RuleDefinition {
  id: string;
  name: string;
  category: string;
  target: RuleTarget;
  property: string;
  operator: 'equals' | 'not_equals' | 'required' | 'forbidden' | 'matches_regex' | 'maximum' | 'minimum' | 'sequence';
  expected?: string | number | boolean;
  unit?: string;
  severity: Severity;
  autoFixable: boolean;
  implemented: boolean;
  evidence: string;
  note?: string;
  conflict?: ConflictAlternative[];
  /** Optional presentation hint for rules that require human confirmation. */
  reviewMode?: 'manual' | 'submission' | 'technical';
  assessmentRole?: 'compliance' | 'consistency';
  appliesWhen?: RuleApplicability;
}

export interface FeatureCondition {
  feature: string;
  equals?: boolean | number | string;
  greaterThan?: number;
  greaterThanOrEqual?: number;
}

export type RuleApplicability =
  | FeatureCondition
  | { all: FeatureCondition[] }
  | { any: FeatureCondition[] };

export interface StrategyDefinition {
  id: string;
  name: string;
  version: string;
  source: string;
  rules: RuleDefinition[];
  /** Built-in strategies are immutable; imported strategies are local copies. */
  origin?: 'built-in' | 'imported';
  sourceTemplateName?: string;
  importedAt?: string;
  baseStrategyId?: string;
  approvedChangeCount?: number;
}

export interface ReferenceStyleSnapshot {
  name: string | null;
  type: string | null;
  basedOn: string | null;
  nextParagraphStyle: string | null;
  fontName: string | null;
  fontSize: number | null;
  alignment: string | null;
  leftIndent: number | null;
  hangingIndent: number | null;
}

export type ReferenceNumberingKind = 'automatic' | 'manual' | 'mixed' | 'none' | 'unknown';

export interface ReferenceNumberingSnapshot {
  kind: ReferenceNumberingKind;
  level: number | null;
  format: string | null;
  suffix: string | null;
  separator?: string | null;
  alignment?: string | null;
  source?: 'direct' | 'style' | 'none';
  startAt: number | null;
  numberPositionPt: number | null;
  textPositionPt: number | null;
}

export interface ParagraphSnapshot {
  index: number;
  text: string;
  style: string;
  styleBuiltIn: string;
  alignment: string;
  firstLineIndent: number | null;
  leftIndent: number | null;
  rightIndent: number | null;
  spaceBefore: number | null;
  spaceAfter: number | null;
  lineSpacing: number | null;
  /** Normalized line-spacing multiple where 1.0 means single spacing. */
  lineSpacingMultiple: number | null;
  lineSpacingRule: string | null;
  /** Raw OOXML w:line value in twentieths of a point, when available. */
  lineSpacingValue?: number | null;
  /** OOXML resolution provenance for developer diagnostics. */
  formattingSources?: Record<string, string>;
  /** 0 means outside a table; positive values are nested table-cell paragraphs. */
  tableNestingLevel: number;
  /** Top-level OOXML table containing this paragraph, when applicable. */
  containingTableIndex?: number | null;
  /** Logical row within the containing table, when applicable. */
  containingTableRowIndex?: number | null;
  fontName: string | null;
  fontSize: number | null;
  bold: boolean | null;
  italic: boolean | null;
  /** Live italic state for association-body text, excluding its marker. */
  authorAssociationBodyItalic?: boolean | null;
  /** Effective Word All Caps formatting for the paragraph text. */
  allCaps: boolean | null;
  isListItem: boolean;
  listString: string | null;
  /** Run-level formatting for known mixed-format prefixes. */
  abstractPrefixBold: boolean | null;
  /** True only when the paragraph's remaining text is entirely bold. */
  abstractBodyBold: boolean | null;
  /** True when the remaining abstract text contains mixed bold/non-bold runs. */
  abstractBodyMixed: boolean | null;
  figurePrefixBold: boolean | null;
  figurePrefixAllCaps: boolean | null;
  tablePrefixBold: boolean | null;
  tablePrefixAllCaps: boolean | null;
  effectiveRuns?: EffectiveRunFormatting[];
  tabStops?: TabStopSnapshot[];
  /** True when the paragraph contains no meaningful text. Resolved from OOXML. */
  isBlank?: boolean;
  /** Structural details used to decide whether an empty paragraph is safe to remove. */
  blankParagraph?: BlankParagraphInfo;
  /** Number of ordinary manual line breaks before the first visible text. Null means unavailable or unsafe to infer. */
  leadingManualLineBreakCount?: number | null;
  /** Effective Word paragraph-level page break-before setting. */
  hasPageBreakBefore?: boolean;
  /** True when the paragraph contains Word field markup. */
  hasField?: boolean;
  /** Visible-text offsets occupied by Word field results. */
  fieldTextSpans?: FieldTextSpan[];
  /** False when this paragraph's live text did not match its OOXML candidate. */
  ooxmlAligned?: boolean;
  /** Signals that the paragraph participates in positioned/manual layout. */
  layoutSensitive?: boolean;
  layoutSignals?: ParagraphLayoutSignals;
  /** Style-layer metadata used by the Reference formatting rules. */
  referenceStyle?: ReferenceStyleSnapshot;
  /** List-layer metadata kept separate from paragraph-style indentation. */
  referenceNumbering?: ReferenceNumberingSnapshot;
  /** Canonical section families inferred from a probable first-level heading. */
  sectionFamilies?: CanonicalSection[];
  /** Normalized heading text used for section-family matching. */
  sectionNormalizedTitle?: string;
  /** Confidence of the section-title recognition, when available. */
  sectionConfidence?: SectionRecognitionConfidence;
  /** Independent confidence that this paragraph is a first-level heading. */
  heading1Confidence?: SectionRecognitionConfidence;
  classification: Exclude<RuleTarget, 'document'> | 'other';
}

export interface BlankParagraphInfo {
  whitespaceOnly: boolean;
  hasManualBreak: boolean;
  hasExplicitPageBreak: boolean;
  hasLastRenderedPageBreak: boolean;
  hasSectionBreak: boolean;
  hasDrawing: boolean;
  hasField: boolean;
  hasEquation: boolean;
  hasNumbering: boolean;
  hasBookmarkContent: boolean;
  safeToDelete: boolean;
}

export interface FieldTextSpan {
  start: number;
  end: number;
  /** A field-backed table number that may be formatted without changing its code. */
  kind?: 'tableSequence' | 'other';
}

export interface ParagraphLayoutSignals {
  hasDrawing: boolean;
  hasAnchoredDrawing: boolean;
  hasCustomTabs: boolean;
  hasNumbering: boolean;
}

export interface TabStopSnapshot {
  positionPt: number;
  alignment: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'list' | 'unknown';
}

export interface ParagraphOoxmlSnapshot {
  index: number;
  /** Text from the same main-flow paragraph sequence used for index alignment. */
  text?: string;
  lineSpacingMultiple: number | null;
  lineSpacingRule: string | null;
  lineSpacingValue?: number | null;
  lineSpacingSource: 'direct' | 'style' | 'basedOn' | 'docDefaults' | 'default' | 'unavailable';
  fontName?: string | null;
  fontSize?: number | null;
  bold?: boolean | null;
  italic?: boolean | null;
  allCaps?: boolean | null;
  figurePrefixAllCaps?: boolean | null;
  tablePrefixAllCaps?: boolean | null;
  alignment?: string | null;
  firstLineIndent?: number | null;
  leftIndent?: number | null;
  rightIndent?: number | null;
  spaceBefore?: number | null;
  spaceAfter?: number | null;
  formattingSources?: Record<string, string>;
  effectiveRuns?: EffectiveRunFormatting[];
  tabStops?: TabStopSnapshot[];
  hasField?: boolean;
  fieldTextSpans?: FieldTextSpan[];
  layoutSensitive?: boolean;
  layoutSignals?: ParagraphLayoutSignals;
  referenceStyle?: ReferenceStyleSnapshot;
  referenceNumbering?: ReferenceNumberingSnapshot;
  containingTableIndex?: number | null;
  containingTableRowIndex?: number | null;
  isBlank?: boolean;
  blankParagraph?: BlankParagraphInfo;
  /** Number of ordinary manual line breaks before the first visible text. */
  leadingManualLineBreakCount?: number | null;
  hasPageBreakBefore?: boolean;
}

export interface EffectiveRunFormatting {
  text: string;
  fontName: string | null;
  fontSize: number | null;
  bold: boolean | null;
  italic: boolean | null;
  allCaps: boolean | null;
  verticalAlign?: 'superscript' | 'subscript' | null;
  /** Positive Word w:position values, converted from half-points. */
  baselinePositionPt?: number | null;
  sources: Record<string, string>;
}

export interface SectionLayoutSnapshot {
  index: number;
  /** Stable structural signature excluding mutable page size and margins. */
  boundarySignature?: string;
  pageWidthInches?: number;
  pageHeightInches?: number;
  orientation?: 'portrait' | 'landscape';
  marginTopInches?: number;
  marginRightInches?: number;
  marginBottomInches?: number;
  marginLeftInches?: number;
  columns?: number;
}

export interface DocumentLayoutSnapshot {
  sections?: SectionLayoutSnapshot[];
  paragraphOoxml?: ParagraphOoxmlSnapshot[];
  documentBlocks?: DocumentBlockSnapshot[];
  pageWidthInches?: number;
  pageHeightInches?: number;
  marginTopInches?: number;
  marginRightInches?: number;
  marginBottomInches?: number;
  marginLeftInches?: number;
  columns?: number;
  hasHeaderReference?: boolean;
  hasFooterReference?: boolean;
  containsPageNumberField?: boolean;
  ooxmlAvailable: boolean;
  /** True only when the OOXML main-flow paragraph count matches Office.js. */
  paragraphMappingAligned?: boolean;
  ooxmlParser?: 'DOMParser' | 'fallback';
  ooxmlError?: string;
}

export type DocumentBlockKind = 'paragraph' | 'table';

export interface EffectiveTableBorders {
  top: boolean | null;
  bottom: boolean | null;
  left: boolean | null;
  right: boolean | null;
  insideH: boolean | null;
  insideV: boolean | null;
  sources?: Partial<Record<'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV', string>>;
}

export interface DocumentBlockSnapshot {
  blockIndex: number;
  kind: DocumentBlockKind;
  paragraphIndex?: number;
  tableIndex?: number;
  /** True for equation-layout tables that should not be treated as manuscript tables. */
  isEquationLayout?: boolean;
  drawingCount?: number;
  hasInlineDrawing?: boolean;
  hasAnchoredDrawing?: boolean;
  tableStyleId?: string | null;
  effectiveTableBorders?: EffectiveTableBorders;
  tableTopBorder?: boolean | null;
  tableBottomBorder?: boolean | null;
  tableHeaderRule?: boolean | null;
  tableHeaderRowCount?: number | null;
  tableCaptionRowIndex?: number | null;
  tableHeaderStartRowIndex?: number | null;
  tableHeaderEndRowIndex?: number | null;
  tableVerticalRules?: boolean | null;
  tableExtraHorizontalRules?: boolean | null;
  text?: string;
}

export interface TableCaptionTableGapIssue {
  captionParagraphIndex: number;
  tableBlockIndex: number;
  blankParagraphIndices: number[];
  blockedByStructuralParagraph?: boolean;
}

export interface TableCaptionTableGapAnalysis {
  issues: TableCaptionTableGapIssue[];
  applicableCaptionCount: number;
  ignoredCaptionIndices: number[];
  blockedCaptionIndices: number[];
}

export interface CitationOccurrence {
  kind: 'reference' | 'figure' | 'table';
  numbers: number[];
  paragraphIndex: number;
  rawText: string;
  /** Optional exact citation-prefix metadata used by deterministic spelling fixes. */
  start?: number;
  end?: number;
  originalText?: string;
  /** The complete matched citation, including its figure number/list. */
  originalCitationText?: string;
  /** The complete citation after replacing only its Figure/Fig. prefix. */
  correctedCitationText?: string;
  /** True when the matched visible citation overlaps a Word field result. */
  isFieldBacked?: boolean;
  isSentenceStart?: boolean;
  expectedPrefix?: 'Figure' | 'Figures' | 'Fig.' | 'Figs.';
  occurrenceIndex?: number;
}

export interface TextReplacement {
  paragraphIndex: number;
  /** Search the complete citation so singular/plural prefixes cannot collide. */
  searchText: string;
  /** The only text range that may be mutated after the citation is located. */
  originalPrefix: 'Figure' | 'Figures' | 'Fig.' | 'Figs.';
  replacementPrefix: 'Figure' | 'Figures' | 'Fig.' | 'Figs.';
  replacement: string;
  occurrenceIndex: number;
}

export interface DefinedObject {
  kind: 'reference' | 'figure' | 'table';
  number: number;
  paragraphIndex: number;
  containingTableIndex?: number | null;
  containingTableRowIndex?: number | null;
}

export interface ObjectPlacementRelation {
  kind: 'figure' | 'table';
  number: number;
  captionParagraphIndex: number;
  objectBlockIndex: number | null;
  passed: boolean | null;
  confidence: 'high' | 'medium' | 'low';
}

export type ReferenceCitationStyle = 'square_bracket_numeric' | 'unknown';

export interface CrossReferenceSnapshot {
  referenceDefinitions: DefinedObject[];
  referenceCitations: CitationOccurrence[];
  referenceCitationStyle: ReferenceCitationStyle;
  figureDefinitions: DefinedObject[];
  figureCitations: CitationOccurrence[];
  tableDefinitions: DefinedObject[];
  tableCitations: CitationOccurrence[];
  referenceCoverageMissing: number[];
  invalidReferenceCitations: CitationOccurrence[];
  referenceSequenceViolations: CitationOccurrence[];
  figureCoverageMissing: number[];
  tableCoverageMissing: number[];
  figurePlacement: ObjectPlacementRelation[];
  tablePlacement: ObjectPlacementRelation[];
  figurePriorCitationFailures: DefinedObject[];
  figurePriorCitationUnknown: DefinedObject[];
  /** Figure citation occurrences that introduce a new figure out of order. */
  figureCitationSequenceViolations?: CitationOccurrence[];
}

export interface DocumentSnapshot {
  paragraphs: ParagraphSnapshot[];
  layout: DocumentLayoutSnapshot;
  referencesHeadingIndex: number | null;
  abstractIndex: number | null;
  crossReferences?: CrossReferenceSnapshot;
  features?: DocumentFeatures;
}

export interface DocumentFeatures {
  authorCount: number;
  affiliationCount: number;
  emailCount: number;
  hasMultipleAuthors: boolean;
  hasMultipleAffiliations: boolean;
  hasMultipleEmails: boolean;
  hasEmailAssociationMarkers: boolean;
  hasAuthorAssociationMarkers: boolean;
  hasDuplicateAffiliationCandidates: boolean;
  hasMultipleAssociations: boolean;
  figureCount: number;
  tableCount: number;
  figureCaptionCount: number;
  tableCaptionCount: number;
  figureObjectCount: number;
  tableObjectCount: number;
  equationCount: number;
  referenceCount: number;
  hasFigures: boolean;
  hasTables: boolean;
  hasFigureCaptions: boolean;
  hasTableCaptions: boolean;
  hasFigureObjects: boolean;
  hasTableObjects: boolean;
  hasEquations: boolean;
  hasReferences: boolean;
  hasMultiPartFigureCandidates: boolean;
  hasFloatingFigures: boolean;
  hasPdf: boolean;
  submissionMode: boolean;
}

export interface Finding {
  ruleId: string;
  ruleName: string;
  category: string;
  target: RuleTarget;
  status: FindingStatus;
  severity: Severity;
  expected: string;
  detected: string;
  message: string;
  paragraphIndices: number[];
  /** Paragraphs the correction engine may need for context (for example all references for renumbering). */
  fixParagraphIndices?: number[];
  /** Text fingerprints captured when a safe paragraph-targeted fix was validated. */
  fixParagraphTexts?: Record<number, string>;
  /** Exact front-matter transitions approved during validation. */
  frontMatterFixTargets?: FrontMatterTransitionFixTarget[];
  /** Exact prefix replacements for deterministic text corrections. */
  fixTextRanges?: TextReplacement[];
  /** Exact OOXML/Word table indices associated with a structural table finding. */
  fixTableIndices?: number[];
  /** Header-row counts aligned with fixTableIndices for table corrections. */
  fixTableHeaderRowCounts?: number[];
  /** Logical caption/header row metadata aligned with fixTableIndices. */
  fixTableCaptionRowIndices?: Array<number | null>;
  fixTableHeaderStartRowIndices?: Array<number | null>;
  fixTableHeaderEndRowIndices?: Array<number | null>;
  autoFixable: boolean;
  property: string;
  rawExpected?: string | number | boolean;
  evidence: string;
  note?: string;
  reviewKind?: ReviewKind;
  assessmentRole?: 'compliance' | 'consistency';
}

export interface CorrectionVerification {
  ruleId: string;
  requestedParagraphIndices: number[];
  verifiedParagraphIndices: number[];
  failedParagraphIndices: number[];
}

export interface SkippedFigureCitation {
  paragraphIndex: number;
  citation: string;
}

export interface FigureCitationCorrectionResult {
  applied: number;
  skippedFieldBacked: SkippedFigureCitation[];
}

export interface CorrectionAttempt {
  ruleId: string;
  paragraphIndices: number[];
  attempted: boolean;
  /** Office.js operation completed without throwing. */
  executionSucceeded: boolean;
  /** Set when the correction service itself revalidated the affected rule(s). */
  verificationSucceeded?: boolean;
  /** Set only after the fresh validation pass completes. */
  resolvedAfterValidation?: boolean;
  message?: string;
}

export interface CorrectionResult {
  verifications: CorrectionVerification[];
  attempts: CorrectionAttempt[];
  /** Number of approved rule fixes submitted to the correction engine. */
  requestedFixes: number;
  /** Number of approved rule fixes whose Office.js operation completed. */
  executedFixes: number;
  /** Number of approved rule fixes that threw or failed required verification. */
  executionFailures: number;
  /** Details for figure citation fixes that were safely applied or skipped. */
  figureCitationCorrection?: FigureCitationCorrectionResult;
  /** Populated by the taskpane after fresh validation. */
  resolvedFixes?: number;
  unresolvedFixes?: number;
}

export interface FrontMatterCorrectionOutcome {
  removedBlankIndices: number[];
  alreadyResolvedTargets: number;
}

export interface ValidationSummary {
  findings: Finding[];
  /** Backward-compatible alias for complianceScore. */
  score: number;
  complianceScore: number;
  overallScore: number;
  passed: number;
  failed: number;
  compliancePassed: number;
  complianceFailed: number;
  consistencyPassed: number;
  consistencyFailed: number;
  unresolvedFailures: number;
  fixableFailures: number;
  manualReview: number;
  checklist: number;
  notApplicable: number;
  notEvaluated: number;
  scoredRules: number;
}
