#!/usr/bin/env node
/**
 * End-to-end suite: drives the built extension in a real browser against a live
 * SharePoint tenant, and screenshots every panel.
 *
 * Usage:
 *   export SP_TENANT=contoso.sharepoint.com
 *   export SP_USER=someone@contoso.onmicrosoft.com
 *   export SP_PASS='…'
 *   npm run build && npm run test:e2e
 *
 * Runs headed on a virtual display; it re-execs itself under xvfb-run when no
 * DISPLAY is present, so it works unattended in a container.
 */
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const { config, assertUsable } = require('./config');
const { ensureSignedIn } = require('./auth');
const { seedLibrary } = require('./fixtures');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---------------------------------------------------------------------------
// Virtual display bootstrap
// ---------------------------------------------------------------------------

// The extension must load into a headed browser, so a display is required.
// Re-exec under Xvfb rather than making the caller remember to.
if (!process.env.DISPLAY && !process.env.E2E_XVFB_WRAPPED && process.platform === 'linux') {
  const probe = spawnSync('which', ['xvfb-run']);
  if (probe.status === 0) {
    const res = spawnSync(
      'xvfb-run',
      ['-a', '--server-args=-screen 0 1440x900x24', process.execPath, __filename, ...process.argv.slice(2)],
      { stdio: 'inherit', env: { ...process.env, E2E_XVFB_WRAPPED: '1' } }
    );
    process.exit(res.status ?? 1);
  }
  console.error('No DISPLAY and xvfb-run is not installed. Install xvfb or set DISPLAY.');
  process.exit(1);
}

const { chromium } = require('playwright-core');

// ---------------------------------------------------------------------------
// Test harness
// ---------------------------------------------------------------------------

const results = [];
let shotIndex = 0;

async function screenshot(page, name) {
  shotIndex++;
  const file = path.join(config.artifactDir, `${String(shotIndex).padStart(2, '0')}-${name}.png`);
  await page.screenshot({ path: file });
  return file;
}

/** Run one scenario, recording pass/fail without aborting the whole suite. */
async function scenario(name, fn) {
  process.stdout.write(`• ${name} … `);
  try {
    const detail = await fn();
    results.push({ name, ok: true, detail });
    console.log(`ok${detail ? ` (${detail})` : ''}`);
  } catch (err) {
    results.push({ name, ok: false, detail: err.message });
    console.log(`FAIL — ${err.message}`);
  }
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

/**
 * Press an extension shortcut and wait for its panel.
 * Content-script shortcuts are page-level listeners, so synthetic key events
 * reach them — unlike chrome.commands entries, which cannot be triggered here.
 */
async function pressAndExpect(page, key, selector, { timeout = 8000 } = {}) {
  // Dispatch to <body> rather than clicking to focus: viewport coordinates near
  // the origin land on SharePoint's app-launcher button, which opens a flyout
  // over the panel under test. The content script listens on `document`, so a
  // key pressed on body reaches it by bubbling.
  await page.locator('body').press(key);
  await page.locator(selector).first().waitFor({ state: 'visible', timeout });
}

async function dismiss(page) {
  await page.keyboard.press('Escape').catch(() => {});
  await sleep(800);
}

/**
 * Close any open inspector panel via its own close button.
 *
 * Escape does not close the inspectors, so leaving one open means the next
 * Alt+<key> press toggles it shut instead of opening it — the panel is then
 * present in the DOM but hidden, and waiting for it to become visible times
 * out. Always return to a known-closed state between scenarios.
 */
async function closePanels(page) {
  for (const id of ['#sp-inspector-close', '#sp-perm-inspector-close']) {
    const btn = page.locator(id);
    if ((await btn.count()) && (await btn.first().isVisible())) {
      await btn.first().click().catch(() => {});
      await sleep(600);
    }
  }
  await page.keyboard.press('Escape').catch(() => {});
  await sleep(600);
}

/**
 * Wait for a modal to stop re-rendering before interacting with it.
 *
 * `show()` appends the overlay synchronously but clears the search box on a
 * 100ms timer, and the async site-lists fetch triggers a later setResults().
 * Both discard anything typed in the meantime, so tests must let the row count
 * go quiet first — otherwise they race the extension and flake.
 */
async function settleModal(page, { quietFor = 2, interval = 700, max = 20 } = {}) {
  const rows = page.locator('.spqn-result-item');
  let last = -1;
  let stable = 0;
  for (let i = 0; i < max; i++) {
    await sleep(interval);
    const n = await rows.count();
    stable = n === last && n > 0 ? stable + 1 : 0;
    last = n;
    if (stable >= quietFor) return n;
  }
  return last;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

(async () => {
  assertUsable();
  fs.mkdirSync(config.artifactDir, { recursive: true });
  fs.mkdirSync(config.profileDir, { recursive: true });

  console.log(`tenant    ${config.tenant}`);
  console.log(`extension ${config.extensionDir}`);
  console.log(`chromium  ${config.chromium}`);
  console.log(`artifacts ${config.artifactDir}\n`);

  const context = await chromium.launchPersistentContext(config.profileDir, {
    executablePath: config.chromium,
    headless: config.headless,
    viewport: { width: 1440, height: 900 },
    args: [
      '--no-sandbox',
      `--disable-extensions-except=${config.extensionDir}`,
      `--load-extension=${config.extensionDir}`,
    ],
  });

  let failures = 0;
  try {
    const worker =
      context.serviceWorkers()[0] ||
      (await context.waitForEvent('serviceworker', { timeout: 20000 }).catch(() => null));
    assert(worker, 'extension service worker never started — extension failed to load');
    console.log(`extension id ${new URL(worker.url()).host}\n`);

    const page = await context.newPage();
    const consoleLog = [];
    page.on('console', (m) => consoleLog.push(m.text()));

    // --- sign in ---------------------------------------------------------
    const session = await ensureSignedIn(page, {
      onScreenshot: (p, n) => screenshot(p, n),
    });
    assert(session.ok, `sign-in failed: ${JSON.stringify(session)}`);
    console.log(`signed in — site "${session.title}"\n`);

    // --- seed --------------------------------------------------------------
    const items = await seedLibrary(page);
    const firstItem = items[0];
    console.log('');

    const listView = `https://${config.tenant}${config.libraryPath}/Forms/AllItems.aspx`;

    // --- scenarios ---------------------------------------------------------
    // Release builds strip console.log/debug/info (webpack.prod.js terser
    // pure_funcs), so injection cannot be proved from the log. Prove it the way
    // a user would notice: the content script answers the service worker and
    // reports its modal initialised.
    await scenario('content script injects and initialises', async () => {
      await page.goto(listView, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await sleep(config.slowNav);
      await page.bringToFront();

      const reply = await worker.evaluate(async () => {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        return chrome.tabs.sendMessage(tab.id, { action: 'toggle-modal' });
      });
      assert(reply && reply.success, `content script did not respond: ${JSON.stringify(reply)}`);

      await screenshot(page, 'list-view');
      await dismiss(page);
      return 'responded to toggle-modal';
    });

    await scenario('QuickNav modal (Ctrl+K)', async () => {
      await pressAndExpect(page, 'Control+k', '.spqn-modal');
      // Rows populate after the modal paints, so let it settle rather than
      // counting immediately and silently reporting zero.
      const count = await settleModal(page);
      assert(count > 0, 'modal opened but rendered no navigation links');
      await screenshot(page, 'quicknav');
      await dismiss(page);
      return `${count} result rows`;
    });

    await scenario('QuickNav filters as you type', async () => {
      await pressAndExpect(page, 'Control+k', '.spqn-modal');
      const all = await settleModal(page);
      const rows = page.locator('.spqn-result-item .spqn-result-title');
      assert(all > 0, 'modal never rendered any rows');

      await page.locator('.spqn-search-input').fill('recycle');
      await sleep(1500);
      const titles = await rows.allInnerTexts();
      assert(titles.length > 0, 'query "recycle" matched nothing');
      assert(titles.length < all, `filter did not narrow results (${all} → ${titles.length})`);
      assert(
        titles.every((t) => /recycle/i.test(t)),
        `non-matching rows survived the filter: ${titles.join(', ')}`
      );

      // and a query that cannot match should surface the empty state
      await page.locator('.spqn-search-input').fill('zzzznomatch');
      await sleep(1500);
      assert((await rows.count()) === 0, 'rows still shown for an impossible query');
      const empty = await page.locator('.spqn-state-container').innerText();
      assert(/no results/i.test(empty), `unexpected empty-state text: ${empty}`);

      await dismiss(page);
      return `${all} → ${titles.length} on "recycle", empty state on no match`;
    });

    await scenario('QuickNav navigates on Enter', async () => {
      await pressAndExpect(page, 'Control+k', '.spqn-modal');
      await settleModal(page);
      await page.locator('.spqn-search-input').fill('Site Contents');
      await sleep(1500);
      const before = page.url();
      await page.keyboard.press('Enter');
      await sleep(6000);
      const after = page.url();
      assert(after !== before, 'Enter did not navigate');
      assert(/viewlsts\.aspx/i.test(after), `unexpected destination: ${after}`);
      assert((await page.locator('.spqn-modal').count()) === 0, 'modal stayed open after selecting');

      // restore the list view for the scenarios that follow
      await page.goto(listView, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await sleep(config.slowNav);
      return 'reached viewlsts.aspx';
    });

    // The Lists modal is bound to Alt+O through chrome.commands, which is a
    // browser-level binding — synthetic page key events cannot reach it. Drive
    // the same message the service worker sends, so the content-script handler
    // is still covered.
    await scenario('Lists & Libraries modal (toggle-lists message)', async () => {
      await page.bringToFront();
      await worker.evaluate(async () => {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        await chrome.tabs.sendMessage(tab.id, { action: 'toggle-lists' });
      });
      await page.locator('.spqn-modal').first().waitFor({ state: 'visible', timeout: 8000 });
      await screenshot(page, 'lists-modal');
      await dismiss(page);
    });

    await scenario('Column inspector (Alt+C)', async () => {
      await pressAndExpect(page, 'Alt+c', '#sp-column-inspector');
      const columnCards = page.locator('#sp-inspector-columns > *');
      await columnCards.first().waitFor({ state: 'visible', timeout: 10000 });
      const columns = await columnCards.count();
      assert(columns > 0, 'no columns rendered');

      // The list-info block carries the List ID / template metadata.
      const info = await page.locator('#sp-inspector-list-info').innerText();
      assert(/List ID/i.test(info), 'list metadata block missing List ID');
      await screenshot(page, 'column-inspector');
      await closePanels(page);
      return `${columns} column cards, list metadata present`;
    });

    await scenario('Column inspector item values (Alt+C with ?ID)', async () => {
      await page.goto(`${listView}?ID=${firstItem.id}`, {
        waitUntil: 'domcontentloaded',
        timeout: 60000,
      });
      await sleep(config.slowNav);
      await pressAndExpect(page, 'Alt+c', '#sp-column-inspector');
      await sleep(4000);

      // Assert the rendered outcome rather than a log line: release builds
      // strip console.log, and the value block is what the user actually sees.
      const body = await page.locator('#sp-column-inspector').innerText();
      assert(/CURRENT VALUE/i.test(body), 'no current-value block rendered');
      assert(
        new RegExp(`ITEM #${firstItem.id}\\b`, 'i').test(body),
        `value block does not identify item #${firstItem.id}`
      );
      await screenshot(page, 'column-inspector-item-values');
      await closePanels(page);
      return `item #${firstItem.id} values rendered`;
    });

    await scenario('Column inspector reveals hidden columns', async () => {
      await pressAndExpect(page, 'Alt+c', '#sp-column-inspector');
      const cards = page.locator('#sp-inspector-columns > *');
      await cards.first().waitFor({ state: 'visible', timeout: 10000 });
      const shown = await cards.count();

      await page.locator('#sp-inspector-show-hidden').check();
      await sleep(2500);
      const withHidden = await cards.count();
      assert(withHidden > shown, `"Show hidden" did not add columns (${shown} → ${withHidden})`);

      await page.locator('#sp-inspector-search').fill('modified');
      await sleep(2000);
      const filtered = await cards.count();
      assert(filtered > 0, 'filtering by "modified" matched no columns');
      assert(filtered < withHidden, `filter did not narrow (${withHidden} → ${filtered})`);

      await closePanels(page);
      return `${shown} visible → ${withHidden} with hidden → ${filtered} matching "modified"`;
    });

    await scenario('Object inspector (Alt+I)', async () => {
      await pressAndExpect(page, 'Alt+i', '.sp-inspector-wrapper');
      await sleep(3000);
      await screenshot(page, 'object-inspector');
      await closePanels(page);
    });

    await scenario('Permission inspector (Alt+P)', async () => {
      await pressAndExpect(page, 'Alt+p', '#sp-permission-inspector');
      await sleep(4000);
      await screenshot(page, 'permission-inspector');
      await closePanels(page);
    });

    await scenario('content script stays off non-SharePoint pages', async () => {
      const offsite = await context.newPage();
      const offsiteLog = [];
      offsite.on('console', (m) => offsiteLog.push(m.text()));
      try {
        await offsite.goto('https://example.com/', {
          waitUntil: 'domcontentloaded',
          timeout: 60000,
        });
        await sleep(4000);
        assert(
          !offsiteLog.some((l) => /Content Script loaded/.test(l)),
          'content script injected on a non-SharePoint origin'
        );
        assert((await offsite.locator('.spqn-modal').count()) === 0, 'modal present off-tenant');
      } finally {
        await offsite.close();
      }
    });

    // Guards a whole class of defect: a template referencing a placeholder the
    // engine does not implement resolves to null, so the link is dropped from
    // the results with nothing but a console warning. Found exactly this with
    // {tenantAdminUrl}, which silently removed "SharePoint Admin Center".
    await scenario('every link template resolves', async () => {
      const unknown = [...new Set(consoleLog.filter((l) => /Unknown placeholder/.test(l)))];
      assert(
        unknown.length === 0,
        `${unknown.length} unsupported placeholder(s) — affected links are silently dropped: ${unknown.join(' | ')}`
      );
    });

    await scenario('no uncaught page errors', async () => {
      const errors = consoleLog.filter((l) => /Uncaught|TypeError:|ReferenceError:/.test(l));
      assert(errors.length === 0, `${errors.length} error(s): ${errors.slice(0, 3).join(' | ')}`);
    });

    fs.writeFileSync(
      path.join(config.artifactDir, 'console.log'),
      consoleLog.join('\n'),
      'utf8'
    );
  } finally {
    await context.close();
  }

  // --- summary -------------------------------------------------------------
  failures = results.filter((r) => !r.ok).length;
  console.log('\n' + '─'.repeat(58));
  for (const r of results) {
    console.log(`${r.ok ? 'PASS' : 'FAIL'}  ${r.name}${r.detail ? ` — ${r.detail}` : ''}`);
  }
  console.log('─'.repeat(58));
  console.log(`${results.length - failures}/${results.length} passed · artifacts in ${config.artifactDir}`);

  process.exit(failures ? 1 : 0);
})().catch((err) => {
  console.error(`\nsuite aborted: ${err.message}`);
  process.exit(1);
});
