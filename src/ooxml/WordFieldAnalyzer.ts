export type WordFieldKind = 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES' | 'OTHER';

export interface WordField {
  kind: WordFieldKind;
  instruction: string;
  fieldType: 'simple' | 'complex';
  startIndex: number;
  endIndex: number;
}

function decodeXmlEntities(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)));
}

function normalizeFieldInstruction(instruction: string): string {
  return decodeXmlEntities(instruction)
    .replace(/\s+/g, ' ')
    .trim();
}

export function classifyFieldInstruction(instruction: string): WordFieldKind {
  const command = normalizeFieldInstruction(instruction).match(/^([A-Za-z]+)/)?.[1]?.toUpperCase();
  switch (command) {
    case 'PAGE': return 'PAGE';
    case 'NUMPAGES': return 'NUMPAGES';
    case 'SECTIONPAGES': return 'SECTIONPAGES';
    default: return 'OTHER';
  }
}

/** True only for a field that generates a table sequence number. */
export function isTableSequenceFieldInstruction(instruction: string): boolean {
  return /^SEQ\s+TABLE(?:\s|$)/i.test(normalizeFieldInstruction(instruction));
}

function localName(tag: string): string {
  return tag.match(/^<\/?(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)\b/i)?.[1]?.toLowerCase() || '';
}

function attribute(tag: string, name: string): string {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return tag.match(new RegExp(`(?:^|\\s)(?:[A-Za-z_][\\w.-]*:)?${escaped}\\s*=\\s*["']([^"']*)["']`, 'i'))?.[1] || '';
}

function parseSimpleFields(xml: string): WordField[] {
  const fields: WordField[] = [];
  const pattern = /<(?:[A-Za-z_][\w.-]*:)?fldSimple\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?fldSimple\s*>)/gi;
  for (const match of xml.matchAll(pattern)) {
    const block = match[0];
    const startIndex = match.index ?? 0;
    const openingTag = block.match(/^<(?:[A-Za-z_][\w.-]*:)?fldSimple\b[^>]*>/i)?.[0] || block;
    const instruction = decodeXmlEntities(attribute(openingTag, 'instr'));
    fields.push({
      kind: classifyFieldInstruction(instruction),
      instruction,
      fieldType: 'simple',
      startIndex,
      endIndex: startIndex + block.length
    });
  }
  return fields;
}

function parseComplexFields(xml: string): WordField[] {
  const fields: WordField[] = [];
  const active: Array<{ startIndex: number; instructionParts: string[] }> = [];
  // Keep fldChar and instrText tokenization separate. A generic token pattern
  // can greedily consume the slash of a self-closing fldChar and cause the
  // begin marker or the first instruction fragment to be skipped.
  const tokenPattern = /<(?:[A-Za-z_][\w.-]*:)?fldChar\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?fldChar\s*>)|<(?:[A-Za-z_][\w.-]*:)?instrText\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?instrText\s*>/gi;

  for (const match of xml.matchAll(tokenPattern)) {
    const token = match[0];
    const tokenStart = match.index ?? 0;
    const name = localName(token);

    if (name === 'instrtext') {
      if (active.length) {
        const openEnd = token.indexOf('>');
        const closeStart = token.lastIndexOf('</');
        const content = openEnd >= 0 && closeStart > openEnd
          ? token.slice(openEnd + 1, closeStart)
          : '';
        active[active.length - 1].instructionParts.push(decodeXmlEntities(content));
      }
      continue;
    }

    const fieldType = attribute(token, 'fldCharType').toLowerCase();
    if (fieldType === 'begin') {
      active.push({ startIndex: tokenStart, instructionParts: [] });
    } else if (fieldType === 'end' && active.length) {
      const current = active.pop()!;
      const instruction = current.instructionParts.join('');
      fields.push({
        kind: classifyFieldInstruction(instruction),
        instruction,
        fieldType: 'complex',
        startIndex: current.startIndex,
        endIndex: tokenStart + token.length
      });
    }
  }

  return fields;
}

export function findWordFields(xml: string): WordField[] {
  const simple = parseSimpleFields(xml);
  const simpleRanges = simple.map(field => [field.startIndex, field.endIndex] as const);
  const complex = parseComplexFields(xml).filter(field =>
    !simpleRanges.some(([start, end]) => field.startIndex >= start && field.endIndex <= end)
  );
  return [...simple, ...complex].sort((a, b) => a.startIndex - b.startIndex);
}

export function findPageNumberFields(xml: string): WordField[] {
  return findWordFields(xml).filter(field => field.kind !== 'OTHER');
}

export function hasPageNumberField(xml: string): boolean {
  return findPageNumberFields(xml).length > 0;
}
