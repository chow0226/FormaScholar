import { DocumentSnapshot } from './models';

/**
 * Fingerprint the document identity relevant to index-based corrections.
 * Formatting changes are intentionally ignored; text/table structure changes
 * invalidate paragraph and table indices and require revalidation.
 */
export function documentFingerprint(snapshot: DocumentSnapshot): string {
  const normalizedParagraphs = snapshot.paragraphs
    .map(paragraph => [
      paragraph.classification,
      paragraph.containingTableIndex ?? '-',
      paragraph.containingTableRowIndex ?? '-',
      paragraph.hasField === true ? 'field' : 'no-field',
      (paragraph.fieldTextSpans || [])
        .map(span => `${span.start}-${span.end}-${span.kind ?? 'other'}`)
        .join(',') || '-',
      paragraph.text.normalize('NFKC').replace(/\s+/g, ' ').trim()
    ].join('|'))
    .join('\u001f');
  const blocks = (snapshot.layout.documentBlocks || [])
    .map(block => [
      block.kind,
      block.tableIndex ?? '-',
      block.paragraphIndex ?? '-',
      block.drawingCount ?? '-',
      block.tableCaptionRowIndex ?? '-',
      block.tableHeaderStartRowIndex ?? '-',
      block.tableHeaderEndRowIndex ?? '-'
    ].join(':'))
    .join(';');
  return `${snapshot.paragraphs.length}|${normalizedParagraphs}|${blocks}`;
}
