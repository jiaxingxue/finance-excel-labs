// A rehype plugin that marks cell references in rendered lesson markdown (GR-11). A reference
// becomes `<a data-cell-ref="BvA!C6">`; LessonPanel renders those as buttons that select the cell.
// The rules for what counts as a reference live in cellRefs.ts.

import { findTextRefs, formatRange, resolveCodeRef, type LinkScope } from './cellRefs.ts';

// Just the parts of the hast tree this plugin touches (https://github.com/syntax-tree/hast).
interface HastText {
  type: 'text';
  value: string;
}
interface HastElement {
  type: 'element';
  tagName: string;
  properties: Record<string, unknown>;
  children: HastNode[];
}
type HastNode = HastText | HastElement | { type: string; children?: HastNode[] };
interface HastParent {
  children: HastNode[];
}

/** Elements whose text is never scanned for plain-text references. */
const SKIP = new Set(['pre', 'a', 'code']);

export function rehypeCellLinks(scope: LinkScope) {
  return (tree: HastParent) => visit(tree, scope);
}

function visit(parent: HastParent, scope: LinkScope): void {
  const out: HastNode[] = [];
  for (const node of parent.children) {
    if (node.type === 'text') {
      out.push(...splitText((node as HastText).value, scope));
    } else if (node.type === 'element') {
      const el = node as HastElement;
      const ref = el.tagName === 'code' ? codeRef(el, scope) : null;
      if (ref) out.push(link(ref, [el]));
      else {
        if (!SKIP.has(el.tagName)) visit(el, scope);
        out.push(el);
      }
    } else {
      out.push(node);
    }
  }
  parent.children = out;
}

function codeRef(code: HastElement, scope: LinkScope): string | null {
  const [only] = code.children;
  if (code.children.length !== 1 || only?.type !== 'text') return null;
  const range = resolveCodeRef((only as HastText).value, scope);
  return range ? formatRange(range) : null;
}

function splitText(value: string, scope: LinkScope): HastNode[] {
  const refs = findTextRefs(value, scope);
  if (refs.length === 0) return [{ type: 'text', value }];
  const nodes: HastNode[] = [];
  let at = 0;
  for (const { index, length, range } of refs) {
    if (index > at) nodes.push({ type: 'text', value: value.slice(at, index) });
    nodes.push(
      link(formatRange(range), [{ type: 'text', value: value.slice(index, index + length) }]),
    );
    at = index + length;
  }
  if (at < value.length) nodes.push({ type: 'text', value: value.slice(at) });
  return nodes;
}

function link(ref: string, children: HastNode[]): HastElement {
  return { type: 'element', tagName: 'a', properties: { dataCellRef: ref }, children };
}
