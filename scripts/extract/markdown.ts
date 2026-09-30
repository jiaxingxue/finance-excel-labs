// Minimal, fence-aware markdown scanning. Enough for the Labs document's structure;
// lines are kept verbatim so lesson text is byte-for-byte what the document says.

export interface Section {
  /** Heading text without the leading hashes. */
  heading: string;
  level: number;
  /** Line index of the heading. */
  start: number;
  /** Line index one past the last line of the section. */
  end: number;
}

export interface Fence {
  lang: string;
  /** Line index of the opening ``` line. */
  start: number;
  /** Line index of the closing ``` line. */
  end: number;
  body: string;
}

// CommonMark allows up to three spaces of indentation (the doc indents one inside a list).
const FENCE = /^ {0,3}```(\S*)\s*$/;
const HEADING = /^(#{1,6}) (.+?)\s*$/;

export function toLines(md: string): string[] {
  return md.replace(/\r\n?/g, '\n').split('\n');
}

/** Headings outside fenced code blocks. */
function headings(lines: string[]): { level: number; text: string; line: number }[] {
  const out: { level: number; text: string; line: number }[] = [];
  let inFence = false;
  lines.forEach((line, i) => {
    if (FENCE.test(line)) {
      inFence = !inFence;
      return;
    }
    if (inFence) return;
    const m = HEADING.exec(line);
    if (m) out.push({ level: m[1]!.length, text: m[2]!, line: i });
  });
  return out;
}

/** Sections at exactly `level`; each runs until the next heading of the same or higher level. */
export function sections(lines: string[], level: number, from = 0, to = lines.length): Section[] {
  const hs = headings(lines).filter((h) => h.line >= from && h.line < to);
  const out: Section[] = [];
  hs.forEach((h, i) => {
    if (h.level !== level) return;
    const next = hs.slice(i + 1).find((n) => n.level <= level);
    out.push({ heading: h.text, level, start: h.line, end: next ? next.line : to });
  });
  return out;
}

export function fences(lines: string[], from = 0, to = lines.length): Fence[] {
  const out: Fence[] = [];
  let open: { lang: string; start: number } | null = null;
  for (let i = from; i < to; i++) {
    const m = FENCE.exec(lines[i]!);
    if (!m) continue;
    if (open) {
      out.push({ ...open, end: i, body: lines.slice(open.start + 1, i).join('\n') });
      open = null;
    } else {
      open = { lang: m[1]!, start: i };
    }
  }
  if (open) throw new Error(`Unclosed code fence at line ${open.start + 1}`);
  return out;
}

/** Find the one section whose heading starts with `prefix`, or throw. */
export function requireSection(all: Section[], prefix: string): Section {
  const found = all.filter((s) => s.heading.startsWith(prefix));
  if (found.length !== 1) {
    throw new Error(
      `Expected exactly one section starting with "${prefix}", found ${found.length}`,
    );
  }
  return found[0]!;
}
