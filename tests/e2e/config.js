/**
 * Shared configuration and environment resolution for the e2e suite.
 *
 * Credentials are read from the environment only — never commit them, and never
 * write them to a file inside the repo.
 */
const fs = require('fs');
const os = require('os');
const path = require('path');

const REPO_ROOT = path.resolve(__dirname, '..', '..');

/**
 * Locate a full Playwright Chromium build.
 *
 * System Google Chrome cannot be used here: Chrome removed the --load-extension
 * switch in M137, and as of M148 the DisableLoadExtensionCommandLineSwitch
 * feature no longer revives it. CDP Extensions.loadUnpacked reports success but
 * installs nothing. Playwright's Chromium still honours --load-extension, so it
 * is the only workable runtime for loading this extension unpacked.
 */
function resolveChromium() {
  if (process.env.E2E_CHROME_PATH) return process.env.E2E_CHROME_PATH;

  const roots = [
    process.env.PLAYWRIGHT_BROWSERS_PATH,
    path.join(os.homedir(), '.cache', 'ms-playwright'),
  ].filter(Boolean);

  for (const root of roots) {
    if (!fs.existsSync(root)) continue;
    for (const entry of fs.readdirSync(root)) {
      if (!entry.startsWith('chromium-')) continue;
      const bin = path.join(root, entry, 'chrome-linux64', 'chrome');
      if (fs.existsSync(bin)) return bin;
      const alt = path.join(root, entry, 'chrome-linux', 'chrome');
      if (fs.existsSync(alt)) return alt;
    }
  }
  return null;
}

const config = {
  repoRoot: REPO_ROOT,
  extensionDir: path.join(REPO_ROOT, 'dist'),
  artifactDir: process.env.E2E_ARTIFACT_DIR || path.join(__dirname, 'artifacts'),

  // Session cookies live outside the repo by default so they cannot be committed.
  profileDir: process.env.E2E_PROFILE_DIR || path.join(os.tmpdir(), 'spqn-e2e-profile'),

  chromium: resolveChromium(),

  tenant: process.env.SP_TENANT,
  user: process.env.SP_USER,
  pass: process.env.SP_PASS,

  // Base32 seed for an enrolled authenticator app. Set only when the account
  // requires MFA and you want unattended runs; treat it as a second password.
  totpSecret: process.env.SP_TOTP_SECRET,

  // The library the fixture documents are seeded into.
  library: process.env.E2E_LIBRARY || 'Documents',
  libraryPath: process.env.E2E_LIBRARY_PATH || '/Shared Documents',

  headless: process.env.E2E_HEADLESS === '1',
  slowNav: Number(process.env.E2E_NAV_WAIT || 9000),
};

/** Fail fast with an actionable message rather than a stack trace deep in Playwright. */
function assertUsable() {
  const missing = [];
  if (!config.tenant) missing.push('SP_TENANT (e.g. contoso.sharepoint.com)');
  if (!config.user) missing.push('SP_USER');
  if (!config.pass) missing.push('SP_PASS');
  if (missing.length) {
    throw new Error(
      `Missing required environment variables:\n  - ${missing.join('\n  - ')}\n\n` +
        'Export them in your shell (not in a file) before running the suite.'
    );
  }
  if (!config.chromium) {
    throw new Error(
      'No Playwright Chromium build found.\n' +
        'Install one with:  npx playwright-core install chromium\n' +
        'or point E2E_CHROME_PATH at an existing chrome binary.\n' +
        'Note: system Google Chrome will NOT work — see the comment in tests/e2e/config.js.'
    );
  }
  if (!fs.existsSync(path.join(config.extensionDir, 'manifest.json'))) {
    throw new Error(`No built extension at ${config.extensionDir}. Run "npm run build" first.`);
  }
}

module.exports = { config, assertUsable, resolveChromium };
