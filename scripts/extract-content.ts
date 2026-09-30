// npm run extract            → regenerate src/data/ from the Labs document
// npm run extract -- --check → exit 1 if src/data/ differs from what the document generates

import { mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync, existsSync } from 'node:fs';
import { dirname, join, relative } from 'node:path';
import { renderFiles } from './extract/content.ts';
import { ContentError } from './extract/validate.ts';

const root = join(import.meta.dirname, '..');
const source = join(root, 'content', 'Excel_Implementation_Labs_Financial_Analysis.md');
const outDir = join(root, 'src', 'data');

function listFiles(dir: string): string[] {
  if (!existsSync(dir)) return [];
  return readdirSync(dir, { recursive: true, withFileTypes: true })
    .filter((d) => d.isFile())
    .map((d) => relative(outDir, join(d.parentPath, d.name)).replaceAll('\\', '/'));
}

function main(): number {
  const check = process.argv.includes('--check');
  let files: Map<string, string>;
  try {
    files = renderFiles(readFileSync(source, 'utf8'));
  } catch (e) {
    if (e instanceof ContentError) {
      console.error(`extract: ${e.message}`);
      return 1;
    }
    throw e;
  }

  const existing = listFiles(outDir);

  if (check) {
    const problems: string[] = [];
    for (const [path, content] of files) {
      const full = join(outDir, path);
      if (!existsSync(full)) problems.push(`missing  src/data/${path}`);
      else if (readFileSync(full, 'utf8') !== content) problems.push(`stale    src/data/${path}`);
    }
    for (const path of existing) {
      if (!files.has(path)) problems.push(`extra    src/data/${path}`);
    }
    if (problems.length > 0) {
      console.error('extract --check: src/data/ is out of date with the Labs document:');
      for (const p of problems) console.error(`  ${p}`);
      console.error('Run `npm run extract` and commit the result.');
      return 1;
    }
    console.log(`extract --check: ${files.size} generated files are up to date.`);
    return 0;
  }

  // src/data/ is 100% generated, so a clean rewrite is safe and removes stale files.
  rmSync(outDir, { recursive: true, force: true });
  for (const [path, content] of files) {
    const full = join(outDir, path);
    mkdirSync(dirname(full), { recursive: true });
    writeFileSync(full, content);
  }
  const manifest = JSON.parse(files.get('manifest.json')!) as { counts: Record<string, number> };
  const c = manifest.counts;
  console.log(
    `extract: wrote ${files.size} files to src/data/ — ${c.sheets} sheets, ${c.cells} cells ` +
      `(${c.formulas} formulas), ${c.assertions} assertions, ${c.experiments} experiments, ` +
      `${c.exercises} pattern cells.`,
  );
  return 0;
}

process.exitCode = main();
