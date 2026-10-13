import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import strategyJson from '../src/strategies/aip-2023.json';
import { StrategyDefinition } from '../src/core/models';
import { isTemplateProposalAcceptable, safeTemplateProposalIds, sharedSpacingValue, TemplateAnalyzer, TemplateAnalysisResult } from '../src/template/TemplateAnalyzer';

const strategy = strategyJson as StrategyDefinition;

function templateXml(): { documentXml: string; stylesXml: string } {
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  return {
    stylesXml: `<w:styles ${ns}>
      <w:style w:type="paragraph" w:styleId="Title">
        <w:name w:val="Title"/><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>
        <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="32"/><w:b/></w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/><w:pPr><w:jc w:val="center"/></w:pPr>
        <w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="24"/><w:b/></w:rPr>
      </w:style>
    </w:styles>`,
    documentXml: `<w:document ${ns}><w:body>
      <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Paper Title</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>INTRODUCTION</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="1"/></w:sectPr>
    </w:body></w:document>`
  };
}

describe('TemplateAnalyzer', () => {
  it('selects only safe changed proposals for bulk strategy review', () => {
    const analysis: TemplateAnalysisResult = {
      fileName: 'AIP-2027.docx',
      components: [],
      candidates: [],
      proposals: [
        { ruleId: 'SAFE_HIGH', ruleName: 'Safe high', component: 'title', property: 'font.size', currentValue: 18, detectedValue: 20, status: 'changed', confidence: 'high' },
        { ruleId: 'SAFE_MEDIUM', ruleName: 'Safe medium', component: 'title', property: 'font.name', currentValue: 'A', detectedValue: 'B', status: 'changed', confidence: 'medium' },
        { ruleId: 'MANUAL', ruleName: 'Manual', component: 'title', property: 'spacing', currentValue: 12, detectedValue: { before: 6, after: 12 }, status: 'changed', confidence: 'manual' },
        { ruleId: 'NEW', ruleName: 'New', component: 'title', property: 'font.italic', detectedValue: true, status: 'new', confidence: 'high' },
        { ruleId: 'UNKNOWN', ruleName: 'Unknown', component: 'title', property: 'font', detectedValue: undefined, status: 'not-determinable', confidence: 'manual' }
      ],
      counts: { unchanged: 0, changed: 3, new: 1, review: 0, notDeterminable: 1 },
      warnings: []
    };

    expect(safeTemplateProposalIds(analysis)).toEqual(['SAFE_HIGH', 'SAFE_MEDIUM']);
  });

  it('extracts measurable styles and compares them without mutating the strategy', async () => {
    const zip = new JSZip();
    const xml = templateXml();
    zip.file('word/document.xml', xml.documentXml);
    zip.file('word/styles.xml', xml.stylesXml);

    const analysis = await TemplateAnalyzer.analyzeBytes(
      await zip.generateAsync({ type: 'uint8array' }),
      'AIP-2027.docx',
      strategy
    );

    const titleSize = analysis.proposals.find(proposal => proposal.ruleId === 'TITLE_SIZE_001');
    expect(titleSize?.status).toBe('changed');
    expect(titleSize?.detectedValue).toBe(16);
    expect(titleSize?.confidence).toBe('high');
    expect(analysis.components.some(component => component.component === 'title')).toBe(true);
    expect(analysis.proposals.some(proposal => proposal.status === 'not-determinable')).toBe(true);
    expect(analysis.proposals.find(proposal => proposal.ruleId === 'PAGE_WIDTH_001')?.detectedValue).toBe(8.5);
    expect(analysis.proposals.find(proposal => proposal.ruleId === 'COLUMN_COUNT_001')?.detectedValue).toBe(1);
    expect(strategy.rules.find(rule => rule.id === 'TITLE_SIZE_001')?.expected).toBe(18);
  });

  it('creates a reviewed strategy copy only for accepted high-confidence changes', async () => {
    const zip = new JSZip();
    const xml = templateXml();
    zip.file('word/document.xml', xml.documentXml);
    zip.file('word/styles.xml', xml.stylesXml);
    const analysis = await TemplateAnalyzer.analyzeBytes(
      await zip.generateAsync({ type: 'uint8array' }),
      'AIP-2027.docx',
      strategy
    );

    const reviewed = TemplateAnalyzer.applyAcceptedChanges(
      strategy,
      analysis,
      new Set(['TITLE_SIZE_001'])
    );

    expect(reviewed.rules.find(rule => rule.id === 'TITLE_SIZE_001')?.expected).toBe(16);
    const reviewedEvidence = reviewed.rules.find(rule => rule.id === 'TITLE_SIZE_001')?.evidence || '';
    expect(reviewedEvidence).toContain('Reviewed template AIP-2027.docx');
    expect(reviewedEvidence).toContain('Detected value: 16 pt');
    expect(reviewedEvidence).not.toContain('Page 1: 18 pt.');
    expect(strategy.rules.find(rule => rule.id === 'TITLE_SIZE_001')?.expected).toBe(18);
  });

  it('uses repeated run-level direct formatting when the named style has no value', async () => {
    const zip = new JSZip();
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    zip.file('word/styles.xml', `<w:styles ${ns}><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style></w:styles>`);
    zip.file('word/document.xml', `<w:document ${ns}><w:body>
      <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="28"/><w:b/></w:rPr><w:t>First Title</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="28"/><w:b/></w:rPr><w:t>Second Title</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="1"/></w:sectPr>
    </w:body></w:document>`);
    const analysis = await TemplateAnalyzer.analyzeBytes(
      await zip.generateAsync({ type: 'uint8array' }),
      'direct-formatting.docx',
      strategy
    );
    const titleSize = analysis.proposals.find(proposal => proposal.ruleId === 'TITLE_SIZE_001');
    expect(titleSize?.detectedValue).toBe(14);
    expect(titleSize?.confidence).toBe('medium');
    expect(titleSize?.evidence?.source).toBe('repeated-examples');
  });

  it('requires manual review when front-matter before/after spacing is asymmetric', async () => {
    const zip = new JSZip();
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    zip.file('word/styles.xml', `<w:styles ${ns}><w:style w:type="paragraph" w:styleId="Abstract"><w:name w:val="Abstract"/><w:pPr><w:spacing w:before="240" w:after="360"/></w:pPr></w:style></w:styles>`);
    zip.file('word/document.xml', `<w:document ${ns}><w:body>
      <w:p><w:pPr><w:pStyle w:val="Abstract"/></w:pPr><w:r><w:t>Abstract. Text</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="1"/></w:sectPr>
    </w:body></w:document>`);
    const analysis = await TemplateAnalyzer.analyzeBytes(
      await zip.generateAsync({ type: 'uint8array' }),
      'asymmetric-spacing.docx',
      strategy
    );
    const spacing = analysis.proposals.find(proposal => proposal.ruleId === 'ABSTRACT_SPACING_001');
    expect(spacing?.detectedValue).toEqual({ before: 12, after: 18 });
    expect(spacing?.confidence).toBe('manual');
    expect(isTemplateProposalAcceptable(spacing!)).toBe(false);
  });

  it('uses the same tolerant shared spacing value when selecting and applying a proposal', () => {
    const proposal = {
      ruleId: 'ABSTRACT_SPACING_001',
      ruleName: 'Abstract spacing',
      component: 'abstract',
      property: 'paragraph.frontMatterSpacing',
      currentValue: 18,
      detectedValue: { before: 13, after: 13.05 },
      status: 'changed' as const,
      confidence: 'high' as const
    };
    const analysis: TemplateAnalysisResult = {
      fileName: 'near-equal-spacing.docx',
      components: [],
      candidates: [],
      proposals: [proposal],
      counts: { unchanged: 0, changed: 1, new: 0, review: 0, notDeterminable: 0 },
      warnings: []
    };

    expect(sharedSpacingValue(proposal.detectedValue)).toBe(13.05);
    expect(isTemplateProposalAcceptable(proposal)).toBe(true);
    const reviewed = TemplateAnalyzer.applyAcceptedChanges(
      strategy,
      analysis,
      new Set(['ABSTRACT_SPACING_001'])
    );
    expect(reviewed.rules.find(rule => rule.id === 'ABSTRACT_SPACING_001')?.expected).toBe(13.05);
  });

  it('uses the tolerant shared spacing rule during full template analysis', async () => {
    const zip = new JSZip();
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    zip.file('word/styles.xml', `<w:styles ${ns}><w:style w:type="paragraph" w:styleId="Abstract"><w:name w:val="Abstract"/><w:pPr><w:spacing w:before="260" w:after="261"/></w:pPr></w:style></w:styles>`);
    zip.file('word/document.xml', `<w:document ${ns}><w:body>
      <w:p><w:pPr><w:pStyle w:val="Abstract"/></w:pPr><w:r><w:t>Abstract. Text</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="1"/></w:sectPr>
    </w:body></w:document>`);

    const analysis = await TemplateAnalyzer.analyzeBytes(
      await zip.generateAsync({ type: 'uint8array' }),
      'near-equal-analysis.docx',
      strategy
    );
    const spacing = analysis.proposals.find(proposal => proposal.ruleId === 'ABSTRACT_SPACING_001');

    expect(spacing?.detectedValue).toEqual({ before: 13, after: 13.05 });
    expect(spacing?.confidence).toBe('high');
    expect(isTemplateProposalAcceptable(spacing!)).toBe(true);
  });

  it('requires manual review when repeated direct formatting conflicts with a named style', async () => {
    const zip = new JSZip();
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    zip.file('word/styles.xml', `<w:styles ${ns}>
      <w:style w:type="paragraph" w:styleId="BaseTitle"><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="36"/></w:rPr></w:style>
      <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="BaseTitle"/><w:rPr><w:b/></w:rPr></w:style>
    </w:styles>`);
    zip.file('word/document.xml', `<w:document ${ns}><w:body>
      <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>First Title</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>Second Title</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1152" w:bottom="1440" w:left="1152"/><w:cols w:num="2"/></w:sectPr>
    </w:body></w:document>`);
    const analysis = await TemplateAnalyzer.analyzeBytes(
      await zip.generateAsync({ type: 'uint8array' }),
      'conflicting-template.docx',
      strategy
    );
    const titleSize = analysis.proposals.find(proposal => proposal.ruleId === 'TITLE_SIZE_001');
    expect(titleSize?.detectedValue).toBe(16);
    expect(titleSize?.confidence).toBe('manual');
    expect(isTemplateProposalAcceptable(titleSize!)).toBe(false);
    expect(analysis.warnings.some(warning => warning.includes('named style'))).toBe(true);
    expect(analysis.proposals.find(proposal => proposal.ruleId === 'COLUMN_COUNT_001')?.detectedValue).toBe(2);
    expect(analysis.proposals.find(proposal => proposal.ruleId === 'MARGIN_TOP_001')?.detectedValue).toBe(0.75);
  });

  it('resolves basedOn style inheritance before comparing template values', async () => {
    const zip = new JSZip();
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    zip.file('word/styles.xml', `<w:styles ${ns}>
      <w:style w:type="paragraph" w:styleId="BaseTitle"><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="36"/></w:rPr></w:style>
      <w:style w:type="paragraph" w:styleId="PaperTitle"><w:name w:val="Paper Title"/><w:basedOn w:val="BaseTitle"/><w:rPr><w:b/></w:rPr></w:style>
    </w:styles>`);
    zip.file('word/document.xml', `<w:document ${ns}><w:body>
      <w:p><w:pPr><w:pStyle w:val="PaperTitle"/></w:pPr><w:r><w:t>Inherited Title</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="1"/></w:sectPr>
    </w:body></w:document>`);
    const analysis = await TemplateAnalyzer.analyzeBytes(
      await zip.generateAsync({ type: 'uint8array' }),
      'inherited-template.docx',
      strategy
    );
    expect(analysis.proposals.find(proposal => proposal.ruleId === 'TITLE_FONT_001')?.detectedValue).toBe('Times New Roman');
    expect(analysis.proposals.find(proposal => proposal.ruleId === 'TITLE_SIZE_001')?.detectedValue).toBe(18);
    expect(analysis.proposals.find(proposal => proposal.ruleId === 'TITLE_BOLD_001')?.detectedValue).toBe(true);
  });

  it('does not let a raised marker run contaminate fallback paragraph formatting', async () => {
    const zip = new JSZip();
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    zip.file('word/styles.xml', `<w:styles ${ns}><w:style w:type="paragraph" w:styleId="AuthorName"><w:name w:val="Author Name"/></w:style></w:styles>`);
    zip.file('word/document.xml', `<w:document ${ns}><w:body>
      <w:p><w:pPr><w:pStyle w:val="AuthorName"/></w:pPr>
        <w:r><w:rPr><w:sz w:val="16"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r>
        <w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>Su-Cheng Haw</w:t></w:r>
      </w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="1"/></w:sectPr>
    </w:body></w:document>`);
    const analysis = await TemplateAnalyzer.analyzeBytes(
      await zip.generateAsync({ type: 'uint8array' }),
      'raised-marker.docx',
      strategy
    );
    const authorSize = analysis.proposals.find(proposal => proposal.ruleId === 'AUTHOR_NAME_SIZE_001');
    expect(authorSize?.detectedValue).toBe(14);
    expect(authorSize?.detectedValue).not.toBe(8);
  });
});
