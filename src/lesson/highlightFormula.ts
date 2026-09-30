// Light syntax highlighting for ```excel code blocks in lessons (LS-1). Those blocks mix formulas
// with short prose labels, so this only picks out strings, function names, and cell references;
// everything else stays plain text.

export type TokenKind = 'string' | 'function' | 'ref' | 'text';

export interface Token {
  kind: TokenKind;
  text: string;
}

const TOKEN = new RegExp(
  [
    '(?<string>"(?:[^"]|"")*")',
    '(?<function>\\b[A-Z][A-Z0-9]*(?:\\.[A-Z0-9]+)*(?=\\())',
    '(?<ref>(?:\\b[A-Za-z][A-Za-z0-9_]*!)?\\$?\\b[A-Z]{1,3}\\$?[0-9]+(?::\\$?[A-Z]{1,3}\\$?[0-9]+)?\\b)',
  ].join('|'),
  'g',
);

export function highlightFormula(source: string): Token[] {
  const tokens: Token[] = [];
  let at = 0;
  for (const m of source.matchAll(TOKEN)) {
    if (m.index > at) tokens.push({ kind: 'text', text: source.slice(at, m.index) });
    const groups = m.groups!;
    const kind: TokenKind =
      groups.string !== undefined ? 'string' : groups.function !== undefined ? 'function' : 'ref';
    tokens.push({ kind, text: m[0] });
    at = m.index + m[0].length;
  }
  if (at < source.length) tokens.push({ kind: 'text', text: source.slice(at) });
  return tokens;
}
