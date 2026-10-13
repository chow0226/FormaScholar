/** Canonical section families recognized by the AIP structure assistant. */
import { romanNumeralToNumber } from './CaptionNumberParser';

export type CanonicalSection =
  | 'introduction'
  | 'background'
  | 'related_work'
  | 'theory_framework'
  | 'problem_definition'
  | 'objectives_questions'
  | 'methodology'
  | 'materials_methods'
  | 'experimental_method'
  | 'proposed_solution'
  | 'system_architecture'
  | 'implementation'
  | 'data'
  | 'evaluation'
  | 'results'
  | 'results_discussion'
  | 'discussion'
  | 'analysis'
  | 'case_study'
  | 'ablation_sensitivity'
  | 'limitations'
  | 'future_work'
  | 'conclusion'
  | 'conclusion_future_work'
  | 'acknowledgments'
  | 'funding'
  | 'declarations'
  | 'data_availability'
  | 'author_contributions'
  | 'ethics'
  | 'references'
  | 'appendix'
  | 'supplementary'
  | 'other';

export type SectionRecognitionConfidence = 'high' | 'medium' | 'low';

export interface SectionRecognition {
  normalizedTitle: string;
  canonicalSections: CanonicalSection[];
  confidence: SectionRecognitionConfidence;
  heading1Confidence: SectionRecognitionConfidence;
  score: number;
}

const SECTION_ALIASES: Record<CanonicalSection, readonly string[]> = {
  introduction: [
    'INTRODUCTION', 'INTRODUCTORY REMARKS', 'OVERVIEW', 'GENERAL INTRODUCTION'
  ],
  background: [
    'BACKGROUND', 'BACKGROUND STUDY', 'RESEARCH BACKGROUND',
    'THEORETICAL BACKGROUND', 'CONCEPTUAL BACKGROUND'
  ],
  related_work: [
    'LITERATURE REVIEW', 'RELATED WORK', 'RELATED WORKS', 'RELATED RESEARCH',
    'PREVIOUS WORK', 'PRIOR WORK', 'STATE OF THE ART'
  ],
  theory_framework: [
    'THEORETICAL FRAMEWORK', 'CONCEPTUAL FRAMEWORK', 'THEORETICAL MODEL',
    'CONCEPTUAL MODEL', 'FRAMEWORK'
  ],
  problem_definition: [
    'PROBLEM STATEMENT', 'PROBLEM FORMULATION', 'PROBLEM DEFINITION',
    'RESEARCH PROBLEM'
  ],
  objectives_questions: [
    'RESEARCH OBJECTIVES', 'OBJECTIVES', 'RESEARCH QUESTIONS', 'HYPOTHESES'
  ],
  methodology: [
    'METHODOLOGY', 'RESEARCH METHODOLOGY', 'METHOD', 'METHODS',
    'RESEARCH METHOD', 'RESEARCH METHODS'
  ],
  materials_methods: [
    'MATERIALS AND METHODS', 'MATERIALS AND METHODOLOGY', 'METHODS AND MATERIALS'
  ],
  experimental_method: [
    'EXPERIMENTAL METHOD', 'EXPERIMENTAL METHODS', 'EXPERIMENTAL METHODOLOGY',
    'EXPERIMENTAL SETUP', 'EXPERIMENTAL DESIGN', 'EXPERIMENTS'
  ],
  proposed_solution: [
    'PROPOSED METHOD', 'PROPOSED METHODOLOGY', 'PROPOSED APPROACH',
    'PROPOSED FRAMEWORK', 'PROPOSED MODEL', 'PROPOSED SYSTEM', 'PROPOSED SOLUTION'
  ],
  system_architecture: [
    'SYSTEM DESIGN', 'SYSTEM ARCHITECTURE', 'MODEL ARCHITECTURE', 'ARCHITECTURE',
    'SYSTEM MODEL', 'NETWORK MODEL'
  ],
  implementation: [
    'IMPLEMENTATION', 'SYSTEM IMPLEMENTATION', 'MODEL IMPLEMENTATION',
    'IMPLEMENTATION DETAILS'
  ],
  data: [
    'DATA', 'DATASET', 'DATASETS', 'DATA COLLECTION', 'DATA ACQUISITION',
    'DATA PREPARATION', 'DATA PREPROCESSING'
  ],
  evaluation: [
    'EVALUATION', 'MODEL EVALUATION', 'SYSTEM EVALUATION',
    'PERFORMANCE EVALUATION', 'EXPERIMENTAL EVALUATION', 'VALIDATION'
  ],
  results: [
    'RESULTS', 'EXPERIMENTAL RESULTS', 'RESEARCH RESULTS', 'FINDINGS',
    'EXPERIMENTAL FINDINGS'
  ],
  results_discussion: [
    'RESULTS AND DISCUSSION', 'RESULTS AND DISCUSSIONS', 'FINDINGS AND DISCUSSION'
  ],
  discussion: [
    'DISCUSSION', 'DISCUSSIONS', 'DISCUSSION OF RESULTS', 'RESULTS DISCUSSION'
  ],
  analysis: [
    'ANALYSIS', 'DATA ANALYSIS', 'PERFORMANCE ANALYSIS', 'COMPARATIVE ANALYSIS',
    'STATISTICAL ANALYSIS'
  ],
  case_study: ['CASE STUDY', 'CASE STUDIES', 'USE CASE', 'USE CASES'],
  ablation_sensitivity: [
    'ABLATION STUDY', 'ABLATION ANALYSIS', 'SENSITIVITY ANALYSIS', 'ROBUSTNESS ANALYSIS'
  ],
  limitations: ['LIMITATIONS', 'STUDY LIMITATIONS', 'LIMITATIONS OF THE STUDY'],
  future_work: [
    'FUTURE WORK', 'FUTURE DIRECTIONS', 'FUTURE RESEARCH',
    'RECOMMENDATIONS FOR FUTURE WORK'
  ],
  conclusion: [
    'CONCLUSION', 'CONCLUSIONS', 'SUMMARY', 'SUMMARY AND CONCLUSION',
    'SUMMARY AND CONCLUSIONS'
  ],
  conclusion_future_work: [
    'CONCLUSION AND FUTURE WORK', 'CONCLUSIONS AND FUTURE WORK',
    'CONCLUSION AND FUTURE DIRECTIONS', 'CONCLUSIONS AND FUTURE DIRECTIONS'
  ],
  acknowledgments: [
    'ACKNOWLEDGMENT', 'ACKNOWLEDGMENTS', 'ACKNOWLEDGEMENT', 'ACKNOWLEDGEMENTS'
  ],
  funding: ['FUNDING', 'FUNDING INFORMATION', 'FUNDING STATEMENT'],
  declarations: [
    'DECLARATIONS', 'DECLARATION', 'CONFLICT OF INTEREST',
    'CONFLICTS OF INTEREST', 'DECLARATION OF INTEREST', 'COMPETING INTERESTS'
  ],
  data_availability: [
    'DATA AVAILABILITY', 'DATA AVAILABILITY STATEMENT',
    'AVAILABILITY OF DATA AND MATERIALS'
  ],
  author_contributions: [
    'AUTHOR CONTRIBUTIONS', "AUTHORS' CONTRIBUTIONS", 'CONTRIBUTIONS'
  ],
  ethics: ['ETHICS STATEMENT', 'ETHICAL APPROVAL', 'ETHICS APPROVAL'],
  references: ['REFERENCES', 'BIBLIOGRAPHY'],
  appendix: ['APPENDIX', 'APPENDICES', 'APPENDIX A', 'APPENDIX B'],
  supplementary: [
    'SUPPLEMENTARY MATERIAL', 'SUPPLEMENTARY MATERIALS', 'SUPPLEMENTAL MATERIAL',
    'SUPPLEMENTARY INFORMATION'
  ],
  other: []
};

const ALIAS_TO_SECTIONS = new Map<string, CanonicalSection[]>();
for (const [section, aliases] of Object.entries(SECTION_ALIASES) as Array<[CanonicalSection, readonly string[]]>) {
  for (const alias of aliases) {
    const existing = ALIAS_TO_SECTIONS.get(alias) || [];
    if (!existing.includes(section)) existing.push(section);
    ALIAS_TO_SECTIONS.set(alias, existing);
  }
}

/** Normalize a probable heading while preserving ordinary words in the title. */
export function normalizeSectionHeading(text: string): string {
  let normalized = text
    .normalize('NFKC')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // Numbering is stripped only when it is a clear heading prefix. Numeric
  // section numbers may omit terminal punctuation (1 Introduction, 1.2
  // Background). Roman prefixes are stripped only when they are canonical;
  // malformed text such as "IC Results" remains visible for review.
  const numericPrefix = /^(?:\d+(?:\.\d+)*)(?:[.)])?\s+/.exec(normalized);
  if (numericPrefix) {
    normalized = normalized.slice(numericPrefix[0].length);
  } else {
    const romanPrefix = /^([IVXLCDM]+)(?:[.)])?\s+/i.exec(normalized);
    if (romanPrefix && romanNumeralToNumber(romanPrefix[1]) !== null) {
      normalized = normalized.slice(romanPrefix[0].length);
    }
  }
  normalized = normalized.replace(/[.:]+$/, '').trim();
  return normalized.replace(/&/g, 'AND').replace(/\s+/g, ' ').toUpperCase();
}

export function canonicalSectionsForHeading(text: string): CanonicalSection[] {
  const normalized = normalizeSectionHeading(text);
  if (/^RESULTS AND (?:ANALYSIS|INTERPRETATION)$/.test(normalized)) {
    return ['results', 'analysis'];
  }
  const exact = ALIAS_TO_SECTIONS.get(normalized);
  if (exact) {
    const expanded = [...exact];
    if (expanded.includes('results_discussion')) expanded.push('results', 'discussion');
    if (expanded.includes('conclusion_future_work')) expanded.push('conclusion', 'future_work');
    return [...new Set(expanded)];
  }

  // Combined headings are intentionally multi-valued so that RESULTS AND
  // DISCUSSION satisfies both conceptual families without requiring a second
  // heading in the manuscript.
  const combined: CanonicalSection[] = [];
  if (/\bRESULTS\b/.test(normalized) && /\bDISCUSSION\b/.test(normalized)) {
    combined.push('results_discussion', 'results', 'discussion');
  }
  if (/\bCONCLUSION(?:S)?\b/.test(normalized) && /\bFUTURE\b/.test(normalized)) {
    combined.push('conclusion_future_work', 'conclusion', 'future_work');
  }
  return [...new Set(combined)];
}

export function recognizeSectionTitle(
  text: string,
  signals: {
    classification?: string;
    style?: string;
    alignment?: string;
    bold?: boolean | null;
    allCaps?: boolean | null;
    fontSize?: number | null;
    isListItem?: boolean;
    tableNestingLevel?: number;
    layoutSensitive?: boolean;
  } = {}
): SectionRecognition {
  const normalizedTitle = normalizeSectionHeading(text);
  const canonicalSections = canonicalSectionsForHeading(text);
  const style = `${signals.style || ''}`.toLowerCase();
  const isHeadingStyle = /heading\s*1|heading\s*2|heading\s*3/.test(style);
  const isHeadingClassification = /^heading[123]$/.test(signals.classification || '');
  const shortEnough = text.trim().length > 0 && text.trim().length <= 120;
  let score = 0;
  if (canonicalSections.length) score += 4;
  if (isHeadingStyle || isHeadingClassification) score += 4;
  if (signals.bold === true) score += 1;
  if (signals.allCaps === true || /^[^a-z]*$/.test(text.trim())) score += 2;
  if (/center/i.test(signals.alignment || '')) score += 1;
  if (!shortEnough) score -= 3;

  const confidence: SectionRecognitionConfidence = score >= 6
    ? 'high'
    : score >= 3
      ? 'medium'
      : 'low';

  let headingScore = 0;
  if (isHeadingClassification && signals.classification === 'heading1') headingScore += 6;
  if (/heading\s*1/.test(style)) headingScore += 6;
  if (canonicalSections.length) headingScore += 1;
  if (signals.bold === true) headingScore += 2;
  if (signals.allCaps === true || /^[^a-z]*$/.test(text.trim())) headingScore += 2;
  if (/center/i.test(signals.alignment || '')) headingScore += 2;
  if (typeof signals.fontSize === 'number' && Math.abs(signals.fontSize - 12) <= 1.5) headingScore += 2;
  if (shortEnough) headingScore += 1;
  if (signals.isListItem === true) headingScore -= 5;
  if ((signals.tableNestingLevel || 0) > 0 || signals.layoutSensitive === true) headingScore -= 10;
  if (/heading\s*[23]/.test(style) || signals.classification === 'heading2' || signals.classification === 'heading3') headingScore -= 10;
  if (text.trim().length > 180) headingScore -= 4;
  const heading1Confidence: SectionRecognitionConfidence = headingScore >= 8
    ? 'high'
    : headingScore >= 5
      ? 'medium'
      : 'low';

  return { normalizedTitle, canonicalSections, confidence, heading1Confidence, score };
}

export function recognizedSectionFamilies(
  paragraphs: Array<{ sectionFamilies?: CanonicalSection[]; sectionConfidence?: SectionRecognitionConfidence }>
): CanonicalSection[] {
  return [...new Set(
    paragraphs
      .filter(p => p.sectionConfidence === 'high' || p.sectionConfidence === 'medium')
      .flatMap(p => p.sectionFamilies || [])
  )];
}

export function sectionFamilyLabel(section: CanonicalSection): string {
  return section.replace(/_/g, ' ');
}
