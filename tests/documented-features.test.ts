/**
 * Documented features must be reachable, and reachable features documented.
 *
 * The README advertised "Recents, favorites and custom links — your own
 * destinations, managed from the options page". Custom links were stored but
 * never shown in the palette; recents were never recorded at all, because
 * selecting a link navigates without calling LinkTracker; and favorites had no
 * control anywhere. The options page had no section for either. Three of the
 * four claims in one bullet were untrue, and every layer passed its own tests.
 *
 * The invariant runs both ways on purpose. Advertising something unbuilt is the
 * bug that happened; wiring something up and leaving the README saying it does
 * not exist is the same defect pointing the other way, and would be silent.
 *
 * What each feature is pinned to matters. "Is the module in the bundle" is
 * necessary but not sufficient: FavoritesManager and LinkTracker are both
 * bundled today, dragged in by a barrel re-export from link-manager, while
 * nothing a user can reach calls either. So recents and favorites are pinned to
 * the component that renders them, and custom links to the call site that was
 * missing.
 */

import * as fs from 'fs';
import * as path from 'path';
import { describe, it, expect } from '@jest/globals';
import { reachableModules } from './helpers/module-graph';

const REPO_ROOT = path.resolve(__dirname, '..');
const read = (rel: string) => fs.readFileSync(path.join(REPO_ROOT, rel), 'utf8');
const README = read('README.md');

/**
 * The feature list only — "## Features" up to the next top-level heading.
 * Elsewhere these words describe the source tree or a roadmap item, which is
 * not a claim about what the extension does.
 */
function featureSection(): string {
  const start = README.indexOf('## Features');
  expect(start).toBeGreaterThan(-1);
  const rest = README.slice(start + '## Features'.length);
  const end = rest.search(/\n## /);
  return end === -1 ? rest : rest.slice(0, end);
}

// quick-access.ts is the only component that renders either list. While it is
// imported by nothing, neither feature exists however much storage code does.
const QUICK_ACCESS = 'src/ui/quick-access.ts';

const RENDERED_FEATURES = [
  // "recycle bin" trips neither; "recently" would, hence the word boundary.
  { name: 'recents', advertised: /\brecents?\b/i },
  { name: 'favorites', advertised: /\bfavou?rites?\b/i },
];

describe('README feature claims match what ships', () => {
  it.each(RENDERED_FEATURES)('$name', ({ name, advertised }) => {
    const claimed = advertised.test(featureSection());
    const shipped = reachableModules().has(QUICK_ACCESS);

    expect({ feature: name, claimed, shipped }).toEqual({
      feature: name,
      claimed: shipped,
      shipped,
    });
  });

  it('custom links: advertised, and the palette actually loads them', () => {
    const claimed = /\bcustom links?\b/i.test(featureSection());
    // The storage module was reachable throughout the bug, because the options
    // page imports it. What was missing is this: the content script reading it.
    const loadedByPalette = /CustomLinksStorage/.test(read('src/content/content-script.ts'));

    expect({ claimed, loadedByPalette }).toEqual({ claimed: true, loadedByPalette: true });
  });
});

describe('module reachability', () => {
  it('reaches the modules behind every shipped inspector', () => {
    const reachable = reachableModules();

    for (const module of [
      'src/ui/column-inspector.ts',
      'src/ui/object-inspector.ts',
      'src/ui/permission-inspector.ts',
      'src/links/fuzzy-search.ts',
      'src/links/link-manager.ts',
    ]) {
      expect([module, reachable.has(module)]).toEqual([module, true]);
    }
  });

  it('resolves the entry points themselves', () => {
    const reachable = reachableModules();

    expect(reachable.has('src/content/content-script.ts')).toBe(true);
    expect(reachable.has('src/options/options.ts')).toBe(true);
    expect(reachable.has('src/background/service-worker.ts')).toBe(true);
  });
});
