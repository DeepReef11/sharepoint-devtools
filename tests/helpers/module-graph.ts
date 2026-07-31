/**
 * Import-graph reachability from the extension's entry points.
 *
 * A module can be present, exported and covered by tests while being imported
 * by nothing that ships — webpack bundles only what the three entries in
 * webpack.common.js pull in, so anything outside that graph is not in the
 * extension at all. Tests passing says nothing about whether a user can reach
 * the code.
 */

import * as fs from 'fs';
import * as path from 'path';

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const SRC = path.join(REPO_ROOT, 'src');

/** Mirrors the `entry` map in webpack.common.js. */
export const ENTRY_POINTS = [
  'content/content-script.ts',
  'options/options.ts',
  'background/service-worker.ts',
];

function resolveImport(fromFile: string, specifier: string): string | null {
  // Bare specifiers are node_modules; only first-party files matter here.
  if (!specifier.startsWith('.')) return null;

  const base = path.resolve(path.dirname(fromFile), specifier);
  for (const candidate of [`${base}.ts`, path.join(base, 'index.ts')]) {
    if (fs.existsSync(candidate)) return candidate;
  }
  return null;
}

/**
 * Every module webpack would pull in, as repo-relative paths.
 */
export function reachableModules(): Set<string> {
  const seen = new Set<string>();

  const visit = (file: string): void => {
    if (seen.has(file)) return;
    seen.add(file);

    const source = fs.readFileSync(file, 'utf8');
    // Covers `import x from './y'`, `import './y'` and `export … from './y'`.
    for (const match of source.matchAll(/(?:from|import)\s+['"]([^'"]+)['"]/g)) {
      const resolved = resolveImport(file, match[1]);
      if (resolved) visit(resolved);
    }
  };

  for (const entry of ENTRY_POINTS) visit(path.join(SRC, entry));

  return new Set([...seen].map((f) => path.relative(REPO_ROOT, f)));
}

/**
 * Whether a repo-relative module is in the shipped bundle.
 */
export function isReachable(relativePath: string): boolean {
  return reachableModules().has(relativePath);
}
