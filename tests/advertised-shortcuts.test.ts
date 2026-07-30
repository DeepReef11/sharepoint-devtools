/**
 * Advertised Shortcut Tests
 *
 * Every shortcut the UI or docs promise must actually be implemented somewhere.
 * There are only two places a shortcut can come from:
 *
 *   page-level  a keydown listener in a content-script module (Ctrl+K, Alt+C,
 *               Alt+I, Alt+P) — these work while a SharePoint page has focus
 *   browser      a `commands` entry in the manifest (Alt+N, Alt+O) — these are
 *               bound by Chrome and reach the extension through the service
 *               worker
 *
 * Ctrl+L was dropped from the second group because it collides with Chrome's
 * address bar, but the popup kept advertising it long after the README was
 * corrected — so the popup promised a key that did nothing. The popup was also
 * missing Alt+P, which had been implemented all along.
 */
import { describe, it, expect } from '@jest/globals';
import * as fs from 'fs';
import * as path from 'path';

const ROOT = path.resolve(__dirname, '..');
const read = (p: string) => fs.readFileSync(path.join(ROOT, p), 'utf8');

/** Files that make a promise to the user about a keystroke. */
const SURFACES = ['public/popup.html', 'README.md'];

/** Chrome-bound shortcuts, from the manifest's `commands`. */
function manifestShortcuts(): string[] {
  const manifest = JSON.parse(read('public/manifest.json'));
  return Object.values<Record<string, any>>(manifest.commands || {})
    .flatMap((cmd) => Object.values<string>(cmd.suggested_key || {}))
    .map(normalise);
}

/** Page-level shortcuts, from keydown handlers across the content script. */
function pageShortcuts(): string[] {
  const found = new Set<string>();
  const dirs = ['src/content', 'src/ui'];
  const files: string[] = [];
  const walk = (d: string) => {
    const full = path.join(ROOT, d);
    if (!fs.existsSync(full)) return;
    for (const e of fs.readdirSync(full, { withFileTypes: true })) {
      if (e.isDirectory()) walk(path.join(d, e.name));
      else if (e.name.endsWith('.ts')) files.push(path.join(d, e.name));
    }
  };
  dirs.forEach(walk);

  for (const f of files) {
    const src = read(f);
    // e.key === 'k' guarded by ctrlKey/metaKey, or by altKey
    for (const m of src.matchAll(/e\.key(?:\.toLowerCase\(\))?\s*===\s*'([a-z])'/g)) {
      const key = m[1].toUpperCase();
      const window = src.slice(Math.max(0, m.index! - 220), m.index! + 40);
      if (/ctrlKey|metaKey/.test(window)) found.add(`CTRL+${key}`);
      if (/altKey|AltGraph/.test(window)) found.add(`ALT+${key}`);
    }
  }
  return [...found];
}

function normalise(s: string): string {
  return s.toUpperCase().replace(/\s+/g, '').replace(/COMMAND|CMD/g, 'CTRL');
}

/** Shortcut-looking strings a surface promises, e.g. "Ctrl+K or Alt+N". */
function advertised(file: string): string[] {
  const text = read(file);
  return [...text.matchAll(/\b((?:Ctrl|Alt|Cmd)\+[A-Za-z])\b/g)].map((m) => normalise(m[1]));
}

describe('advertised keyboard shortcuts', () => {
  const implemented = new Set([...manifestShortcuts(), ...pageShortcuts()]);

  it('finds both kinds of implementation', () => {
    // Guard against the extraction silently returning nothing, which would make
    // every assertion below pass regardless.
    expect(manifestShortcuts().length).toBeGreaterThan(0);
    expect(pageShortcuts()).toContain('CTRL+K');
  });

  it.each(SURFACES)('%s only advertises shortcuts that exist', (file) => {
    const unimplemented = [...new Set(advertised(file))].filter((s) => !implemented.has(s));
    expect(unimplemented).toEqual([]);
  });

  it('does not advertise Ctrl+L, which was removed for colliding with the address bar', () => {
    for (const file of SURFACES) {
      expect(advertised(file)).not.toContain('CTRL+L');
    }
  });

  it('advertises every inspector the content script implements', () => {
    // The popup shipped without Alt+P for some time despite it working.
    const popup = advertised('public/popup.html');
    for (const key of ['ALT+C', 'ALT+I', 'ALT+P']) {
      expect(popup).toContain(key);
    }
  });
});
