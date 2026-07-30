/**
 * Microsoft 365 sign-in against the persistent profile.
 *
 * The profile keeps session cookies, so after the first successful run the
 * login flow is skipped entirely. The flow is driven by what is actually on
 * screen rather than a fixed sequence, because Entra varies the gates it shows
 * (account picker, password, "Stay signed in?", consent).
 */
const { config } = require('./config');
const { totp, secondsRemaining } = require('./totp');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const MAX_GATES = 12;

/**
 * Walk the sign-in gates until the tenant loads.
 * @returns {Promise<boolean>} true if the session is authenticated
 */
async function ensureSignedIn(page, { log = console.log, onScreenshot } = {}) {
  await page.goto(`https://${config.tenant}/`, {
    waitUntil: 'domcontentloaded',
    timeout: 60000,
  });
  await sleep(3000);

  // Fast path: a reused profile is usually already authenticated. Check before
  // touching the login flow — walking it needlessly can trip interstitials
  // (such as the security-defaults MFA nag) that the session does not require.
  const existing = await verifySession(page);
  if (existing.ok) {
    log('  auth: existing session reused');
    return existing;
  }

  let lastAction = null;
  let repeats = 0;
  let blockedByMfa = false;

  for (let i = 0; i < MAX_GATES; i++) {
    await sleep(4000);

    const url = page.url();
    if (!url.includes('login.microsoftonline.com') && !url.includes('/_forms/')) break;

    // Repeating the same action means the gate is not advancing. Fail with
    // something actionable rather than hammering it until MAX_GATES.
    if (repeats >= 2) {
      throw new Error(
        `sign-in stalled: "${lastAction}" repeated ${repeats + 1}x without advancing. ` +
          `Screen reads: ${(await page.evaluate(() => document.body.innerText).catch(() => '')).replace(/\s+/g, ' ').slice(0, 200)}`
      );
    }

    if (onScreenshot) await onScreenshot(page, `auth-${i}`);

    const text = await page.evaluate(() => document.body.innerText).catch(() => '');

    // Entra keeps an off-screen password input (#i0118, ~10x13px,
    // aria-hidden="true") mounted on the *email* screen. Playwright's :visible
    // matches it because the box is non-zero, so filtering on aria-hidden is
    // what actually distinguishes the live field from the decoy — and email is
    // checked first, since it is always the earlier gate.
    const live = (type) =>
      page.locator(`input[type="${type}"]:not([aria-hidden="true"]):visible`).first();
    const email = live('email');
    const password = live('password');
    const tile = page
      .locator('[role="button"], div[data-test-id]')
      .filter({ hasText: new RegExp(escapeRe(config.user), 'i') })
      .first();

    let action;
    if (await email.count()) {
      action = 'account submitted';
      await email.fill(config.user);
      await submit(page);
    } else if (await password.count()) {
      action = 'password submitted';
      await password.fill(config.pass);
      await submit(page);
    } else if (
      config.totpSecret &&
      /Enter the (6-digit )?code|verification code|Enter code/i.test(text) &&
      (await page.locator('input[type="tel"]:visible, input[name*="otc" i]:visible').count())
    ) {
      // TOTP challenge. Key off a real input box — the enrolment screens also
      // contain the phrase "Enter code manually" with nothing to type into.
      action = 'MFA code submitted';
      // Don't submit a code that is about to roll over mid-request.
      if (secondsRemaining() < 6) await sleep((secondsRemaining() + 1) * 1000);
      await page
        .locator('input[type="tel"]:visible, input[name*="otc" i]:visible')
        .first()
        .fill(totp(config.totpSecret));
      await submit(page);
    } else if (/Stay signed in|Rester connect/i.test(text)) {
      action = 'stay-signed-in accepted';
      const yes = page
        .locator('input[type="submit"]:visible, button:visible')
        .filter({ hasText: /^(Yes|Oui)$/i })
        .first();
      if (await yes.count()) await yes.click();
      else await page.locator('input[type="submit"]:visible').last().click();
    } else if (await tile.count()) {
      action = 'account tile picked';
      await tile.click();
    } else if (/Action Required|More information required|multifactor/i.test(text)) {
      // Security defaults nag you toward MFA enrolment before it is mandatory.
      // Defer it while a deferral is still offered; once it is not, say so
      // plainly rather than looping.
      const later = page
        .locator('a:visible, button:visible, input[type="button"]:visible')
        .filter({ hasText: /Skip for now|Ask later|Maybe later|remind me/i })
        .first();
      if (await later.count()) {
        action = 'MFA enrolment deferred';
        await later.click();
      } else {
        // Do not give up here — the nag is often transient, and a direct
        // re-navigation below frequently lands on the tenant anyway. Only
        // report it if the session genuinely turns out to be unusable.
        blockedByMfa = true;
        break;
      }
    } else {
      log(`  auth: unrecognised gate — ${text.replace(/\s+/g, ' ').slice(0, 160)}`);
      break;
    }

    repeats = action === lastAction ? repeats + 1 : 0;
    lastAction = action;
    log(`  auth: ${action}`);
  }

  await sleep(5000);

  // verifySession uses a root-relative fetch, so it must run on the tenant
  // origin. If a gate left us parked on login.microsoftonline.com, going
  // straight to verify would report a meaningless 404 instead of the real
  // problem.
  if (!page.url().startsWith(`https://${config.tenant}`)) {
    await page
      .goto(`https://${config.tenant}/`, { waitUntil: 'domcontentloaded', timeout: 60000 })
      .catch(() => {});
    await sleep(5000);
  }

  const session = await verifySession(page);
  if (!session.ok && blockedByMfa) {
    throw new Error(
      'MFA enrolment is required for this account and could not be deferred. ' +
        'Either turn off security defaults in Entra, or sign in once by hand with ' +
        'E2E_PROFILE_DIR pointed at a persistent directory (see tests/e2e/README.md) ' +
        'so the session cookies can be reused.'
    );
  }
  return session;
}

async function submit(page) {
  await page
    .locator('input[type="submit"]:visible, button[type="submit"]:visible')
    .first()
    .click();
}

/** Prove the session is real by calling the same REST endpoint the extension uses. */
async function verifySession(page) {
  const result = await page.evaluate(async () => {
    try {
      const res = await fetch('/_api/web?$select=Title,Url', {
        headers: { Accept: 'application/json;odata=verbose' },
      });
      const type = res.headers.get('content-type') || '';
      if (!type.includes('json')) return { ok: false, status: res.status };
      const body = await res.json();
      return { ok: res.ok, status: res.status, title: body?.d?.Title };
    } catch (err) {
      return { ok: false, error: err.message };
    }
  });
  return result;
}

function escapeRe(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

module.exports = { ensureSignedIn, verifySession };
