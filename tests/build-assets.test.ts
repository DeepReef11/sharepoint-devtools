/**
 * Build Asset Tests
 *
 * Guards a failure mode that type-checking, linting and unit tests all miss: an
 * HTML file referencing an asset the build never emits. `options.html` linked
 * `options.css`, webpack copied the HTML but not the stylesheet, and the options
 * page shipped completely unstyled — with `.modal { display: none }` absent, both
 * dialogs sat permanently expanded in the page body.
 *
 * Nothing failed. The bundle built, the tests passed, only the rendered page was
 * broken. These tests read the real webpack config and assert that every local
 * asset referenced by the extension's HTML is actually produced.
 */
import { describe, it, expect } from '@jest/globals';
import * as fs from 'fs';
import * as path from 'path';

const ROOT = path.resolve(__dirname, '..');

/** HTML files shipped in the extension, with their source paths. */
const HTML_FILES = ['src/options/options.html', 'public/popup.html'];

const webpackConfig = fs.readFileSync(path.join(ROOT, 'webpack.common.js'), 'utf8');

/** Bundles webpack emits, from its `entry` map: `foo: './x'` → `foo.js`. */
function entryOutputs(): string[] {
  const block = webpackConfig.match(/entry:\s*\{([\s\S]*?)\}/);
  if (!block) return [];
  return [...block[1].matchAll(/(\w[\w-]*)\s*:/g)].map((m) => `${m[1]}.js`);
}

/** Files CopyWebpackPlugin places at the dist root. */
function copiedFiles(): string[] {
  const out: string[] = [];
  for (const m of webpackConfig.matchAll(/from:\s*'([^']+)'\s*,\s*to:\s*'([^']+)'/g)) {
    const [, from, to] = m;
    if (to === '.') {
      // A whole directory: everything inside it lands at the root.
      const dir = path.join(ROOT, from);
      if (fs.existsSync(dir)) out.push(...fs.readdirSync(dir));
    } else {
      out.push(to);
    }
  }
  return out;
}

/** Local (non-absolute, non-data) hrefs and srcs in an HTML file. */
function localRefs(htmlPath: string): string[] {
  const html = fs.readFileSync(path.join(ROOT, htmlPath), 'utf8');
  const refs = [...html.matchAll(/(?:href|src)\s*=\s*"([^"]+)"/g)].map((m) => m[1]);
  return refs.filter((r) => !/^(https?:|data:|mailto:|#|\/\/)/.test(r));
}

describe('build assets', () => {
  const produced = new Set([...entryOutputs(), ...copiedFiles()]);

  it('derives a non-empty set of build outputs', () => {
    // If the config parsing silently breaks, the assertions below become
    // vacuous — so check the parse itself first.
    expect(entryOutputs().length).toBeGreaterThan(0);
    expect(copiedFiles().length).toBeGreaterThan(0);
  });

  it.each(HTML_FILES)('every local asset referenced by %s is produced by the build', (htmlPath) => {
    const missing = localRefs(htmlPath).filter((ref) => {
      const name = ref.replace(/^\.?\//, '').split(/[?#]/)[0];
      return !produced.has(name) && !fs.existsSync(path.join(ROOT, 'public', name));
    });

    expect(missing).toEqual([]);
  });

  it('copies the options stylesheet, not just the options markup', () => {
    // The specific regression: the pair must stay together.
    expect(produced.has('options.html')).toBe(true);
    expect(produced.has('options.css')).toBe(true);
  });

  it('references a stylesheet from the options page at all', () => {
    // If the link tag were dropped instead of the file, the page would break the
    // same way while the assertion above still passed.
    const html = fs.readFileSync(path.join(ROOT, 'src/options/options.html'), 'utf8');
    expect(html).toMatch(/<link[^>]+rel="stylesheet"[^>]+href="options\.css"/);
  });
});
