import { DocumentFeatures, DocumentSnapshot, RuleApplicability, RuleDefinition } from './models';

export interface ApplicabilityResult {
  applies: boolean;
  reason: string;
}

function normalizedAffiliation(text: string): string {
  return text
    .replace(/^\s*\d+(?:\s*[,;]\s*\d+)*\s*/, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export class ApplicabilityAnalyzer {
  static analyze(snapshot: DocumentSnapshot): DocumentFeatures {
    const paragraphs = snapshot.paragraphs;
    const affiliations = paragraphs.filter(p => p.classification === 'authorAffiliation');
    const figures = paragraphs.filter(p => p.classification === 'figureCaption');
    const tables = paragraphs.filter(p => p.classification === 'tableCaption');
    const equations = paragraphs.filter(p => p.classification === 'equation');
    const references = paragraphs.filter(p => p.classification === 'reference');
    const authorParagraphs = paragraphs.filter(p => p.classification === 'authorName');
    const emails = paragraphs.filter(p => p.classification === 'authorEmail');
    const authorAssociationMarkers = authorParagraphs.flatMap(p => p.text.match(/[a-z]\)/gi) || []);
    const authorAndGroups = authorParagraphs.filter(p => /\s+and\s+/i.test(p.text)).length;
    const hasMultipleAuthors = authorParagraphs.length > 1 || authorAssociationMarkers.length > 1 || authorAndGroups > 0;
    const affiliationKeys = affiliations.map(p => normalizedAffiliation(p.text)).filter(Boolean);
    const hasDuplicateAffiliationCandidates = new Set(affiliationKeys).size < affiliationKeys.length;
    const hasAuthorAssociationMarkers = paragraphs.some(p =>
      (p.classification === 'authorAffiliation' && /^\s*\d+(?:\s*[,;]\s*\d+)*\s*/.test(p.text))
      || (p.classification === 'authorEmail' && /^\s*[A-Za-z]\)\s*/.test(p.text))
    );
    const hasEmailAssociationMarkers = emails.some(p => /^\s*[A-Za-z]\)/.test(p.text));
    const captionMentionsSubparts = figures.some(p =>
      /(?:\([a-z]\)|\b[a-z]\))/i.test(p.text)
      || (p.effectiveRuns || []).some(run => /\([a-z]\)/i.test(run.text))
    );
    const blocks = snapshot.layout.documentBlocks || [];
    const multiDrawingFigure = snapshot.crossReferences?.figurePlacement.some(relation => {
      if (relation.confidence === 'low' || relation.objectBlockIndex === null) return false;
      const block = blocks.find(candidate => candidate.blockIndex === relation.objectBlockIndex);
      return (block?.drawingCount || 0) > 1;
    }) || false;
    const authorCount = Math.max(
      authorParagraphs.length,
      authorAssociationMarkers.length,
      authorAndGroups > 0 ? 2 : 0
    );
    const figureObjectCount = blocks.filter(block => (block.drawingCount || 0) > 0).length;
    const tableObjectCount = blocks.filter(block =>
      block.kind === 'table'
      && block.isEquationLayout !== true
      && (block.drawingCount || 0) === 0
    ).length;
    const hasFigureCaptions = figures.length > 0;
    const hasTableCaptions = tables.length > 0;
    const hasFigureObjects = figureObjectCount > 0;
    const hasTableObjects = tableObjectCount > 0;

    return {
      authorCount,
      affiliationCount: affiliations.length,
      emailCount: emails.length,
      hasMultipleAuthors,
      hasMultipleAffiliations: affiliations.length > 1,
      hasMultipleEmails: paragraphs.filter(p => p.classification === 'authorEmail').length > 1,
      hasEmailAssociationMarkers,
      hasAuthorAssociationMarkers,
      hasDuplicateAffiliationCandidates,
      hasMultipleAssociations: hasMultipleAuthors
        && (affiliations.length > 1 || paragraphs.filter(p => p.classification === 'authorEmail').length > 1 || authorAssociationMarkers.length > 1),
      figureCount: figures.length,
      tableCount: tables.length,
      figureCaptionCount: figures.length,
      tableCaptionCount: tables.length,
      figureObjectCount,
      tableObjectCount,
      equationCount: equations.length,
      referenceCount: references.length,
      hasFigures: hasFigureCaptions || hasFigureObjects,
      hasTables: hasTableCaptions || hasTableObjects,
      hasFigureCaptions,
      hasTableCaptions,
      hasFigureObjects,
      hasTableObjects,
      hasEquations: equations.length > 0,
      hasReferences: references.length > 0,
      hasMultiPartFigureCandidates: captionMentionsSubparts || multiDrawingFigure,
      hasFloatingFigures: blocks.some(block => block.hasAnchoredDrawing === true),
      hasPdf: false,
      submissionMode: false
    };
  }

  static check(rule: RuleDefinition, snapshot: DocumentSnapshot): ApplicabilityResult {
    const features = snapshot.features || this.analyze(snapshot);
    const explicit = rule.appliesWhen ? this.evaluateCondition(rule.appliesWhen, features) : null;
    if (explicit !== null) {
      return explicit
        ? { applies: true, reason: 'The configured manuscript feature condition is satisfied.' }
        : { applies: false, reason: 'The configured manuscript feature condition is not satisfied.' };
    }

    const id = rule.id;
    if (id === 'AUTHOR_ASSOCIATION_001') return this.result(features.hasMultipleAssociations, 'Multiple authors with multiple author associations were detected.');
    if (id === 'AUTHOR_SHARED_AFFILIATION_001') return this.result(features.hasMultipleAuthors && (features.hasMultipleAffiliations || features.hasAuthorAssociationMarkers), 'Multiple authors have affiliation associations that require shared-affiliation review.');
    if (id === 'AUTHOR_AFFILIATION_SUPERSCRIPT_001') return this.result(features.hasMultipleAffiliations, 'Multiple affiliations require association markers.');
    if (id === 'AUTHOR_EMAIL_SUPERSCRIPT_001') return this.result(
      features.hasMultipleEmails || features.hasEmailAssociationMarkers,
      'Multiple emails or an email association marker require superscript review.'
    );
    if (id === 'FIGURE_SUBPART_LABELS_001') return this.result(features.hasMultiPartFigureCandidates, 'A multi-part figure candidate was detected.');
    if (id === 'FIGURE_CAPTION_PRESENCE_001') return this.result(features.hasFigureObjects, 'Figure/drawing candidates were detected and require caption-coverage review.');
    if (id === 'TABLE_CAPTION_PRESENCE_001') return this.result(features.hasTableObjects, 'Table candidates were detected and require caption-coverage review.');
    if (id === 'EQUATION_READABLE_001' || id === 'EQUATION_TABS_001') return this.result(features.hasEquations, 'Equation paragraphs were detected.');
    if (id === 'OBJECTS_INLINE_001') return this.result(features.hasFigures || features.hasTables || features.hasEquations, 'Figures, tables, or equations were detected.');
    if (id === 'FIGURE_READABILITY_001') return this.result(
      features.submissionMode && features.hasFigures,
      'Figure readability is reviewed during submission readiness when figures are present.'
    );
    if (id === 'COPYRIGHT_PERMISSION_001') return this.result(
      features.submissionMode,
      'Copyright permission review is relevant during submission readiness.'
    );
    if (id === 'MIN_PAGE_COUNT_001') return this.result(features.submissionMode, 'Page-count review is relevant during submission readiness.');
    // This consistency rule must remain applicable when the References section
    // is missing; that missing section is the condition it is designed to
    // report when in-text citations are present.
    if (id === 'REFERENCE_SECTION_CONSISTENCY_001') {
      return { applies: true, reason: 'Reference-section consistency is evaluated against detected in-text citations.' };
    }
    if (id.startsWith('PDF_')) {
      const needsMath = id.includes('MATH');
      const needsFigure = id.includes('FIGURE');
      return this.result(features.submissionMode && (!needsMath || features.hasEquations) && (!needsFigure || features.hasFigures), 'Submission/PDF review is relevant.');
    }
    // Review-mode metadata takes precedence over broad category prefixes such
    // as REFERENCE_ and FIGURE_. Submission checklist items must not leak into
    // ordinary manuscript validation, while the specific checks above retain
    // their own feature conditions.
    if (rule.reviewMode === 'submission') return this.result(features.submissionMode, 'Submission mode is active.');
    if (id.startsWith('FIGURE_')) return this.result(features.hasFigures, 'Figure content was detected.');
    if (id.startsWith('TABLE_')) return this.result(features.hasTables, 'Table content was detected.');
    if (id.startsWith('REFERENCE_')) return this.result(features.hasReferences, 'Reference content was detected.');
    return { applies: true, reason: 'This review remains relevant by default.' };
  }

  private static result(applies: boolean, reason: string): ApplicabilityResult {
    return { applies, reason: applies ? reason : `Not applicable: ${reason}` };
  }

  private static evaluateCondition(condition: RuleApplicability, features: DocumentFeatures): boolean {
    if ('feature' in condition) return this.evaluateFeature(condition, features);
    if ('all' in condition) return condition.all.every(item => this.evaluateFeature(item, features));
    return condition.any.some(item => this.evaluateFeature(item, features));
  }

  private static evaluateFeature(condition: { feature: string; equals?: boolean | number | string; greaterThan?: number; greaterThanOrEqual?: number }, features: DocumentFeatures): boolean {
    const value = (features as unknown as Record<string, unknown>)[condition.feature];
    if (condition.equals !== undefined) return value === condition.equals;
    if (condition.greaterThan !== undefined) return typeof value === 'number' && value > condition.greaterThan;
    if (condition.greaterThanOrEqual !== undefined) return typeof value === 'number' && value >= condition.greaterThanOrEqual;
    return Boolean(value);
  }
}
