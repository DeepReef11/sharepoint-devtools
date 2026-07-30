# End-to-end suite

Drives the built extension in a real Chromium instance against a live SharePoint
tenant, and screenshots every panel. Complements the Jest unit tests, which run
against jsdom and never load the extension.

```bash
npm run build
export SP_TENANT=contoso.sharepoint.com
export SP_USER=someone@contoso.onmicrosoft.com
export SP_PASS='…'
npm run test:e2e
```

Exits non-zero if any scenario fails. Screenshots and the full page console log
land in `tests/e2e/artifacts/` (gitignored).

## What it covers

| Scenario | Asserts |
| --- | --- |
| content script injects and initialises | answers the service worker's `toggle-modal` with `{success:true}` |
| QuickNav modal (Ctrl+K) | modal opens **and** renders navigation rows |
| QuickNav filters as you type | query narrows rows, every row matches, empty state on no match |
| QuickNav navigates on Enter | selection reaches the resolved URL and closes the modal |
| Lists & Libraries | `toggle-lists` message opens the modal |
| Column inspector (Alt+C) | column cards render, list metadata includes List ID |
| Column inspector item values | `?ID=` triggers the item fetch and a current-value block |
| Column inspector hidden columns | "Show hidden" adds columns; the filter box narrows them |
| Object inspector (Alt+I) | panel opens |
| Permission inspector (Alt+P) | panel opens |
| content script scoping | does **not** inject on a non-SharePoint origin |
| every link template resolves | no `Unknown placeholder` warnings (see below) |
| no uncaught page errors | console is free of `Uncaught` / `TypeError` |

### `every link template resolves`

`LinkTemplateEngine.getPlaceholderValue` implements only the four members of the
`Placeholder` enum. A registry template referencing anything else resolves to
`null`, and the link is dropped from the results with nothing but a
`console.warn` — no error, no visible symptom.

This caught exactly that: `registry.ts` used `{tenantAdminUrl}`, which was absent
from the enum, so `sharepoint-admin-center` ("SharePoint Admin Center") never
appeared in QuickNav — 15 rows instead of 16. Fixed by adding `ServerUrl` and
`TenantAdminUrl` to the enum and the resolver switch. `tests/link-templates.test.ts`
now guards the same invariant statically, so `npm test` catches it without a tenant.

### Release builds strip `console.log`

`webpack.prod.js` passes `console.log`/`debug`/`info` to terser as `pure_funcs`,
so they are gone from `dist/`. `console.warn` and `console.error` survive.

Scenarios therefore assert on rendered DOM or on message round-trips, never on
`console.log` strings — the one exception being `every link template resolves`,
which matches a `console.warn` and so still works. Use `npm run build:dev` if you
need the diagnostic logging while debugging a failure.

### Timing

`Modal.show()` appends the overlay synchronously but clears the search box on a
100 ms timer, and the async site-lists fetch fires a later `setResults()`. Both
discard anything typed in between, so any scenario that types into a modal calls
`settleModal()` first to wait for the row count to go quiet. Without it, tests
race the extension and flake between "filter worked" and "filter ignored".

## Requirements

**Playwright's Chromium — not system Google Chrome.** Chrome removed the
`--load-extension` switch in M137, and as of M148 the
`DisableLoadExtensionCommandLineSwitch` feature no longer revives it. CDP
`Extensions.loadUnpacked` reports success but installs nothing (verified against
Chrome 148: `chrome://extensions` stays empty, with or without Developer mode).
Playwright's Chromium still honours the switch.

```bash
npx playwright-core install chromium
```

**A display.** The extension only loads in a headed browser, so `run.js`
re-execs itself under `xvfb-run` when `DISPLAY` is unset. Install `xvfb`, or
export `DISPLAY` yourself.

## Environment

| Variable | Purpose |
| --- | --- |
| `SP_TENANT` | tenant host, e.g. `contoso.sharepoint.com` |
| `SP_USER` / `SP_PASS` | sign-in credentials |
| `SP_TOTP_SECRET` | base32 seed of an enrolled authenticator, for unattended MFA |
| `E2E_CHROME_PATH` | explicit Chromium binary, bypassing auto-discovery |
| `E2E_PROFILE_DIR` | session profile (default: a temp dir, **outside the repo**) |
| `E2E_ARTIFACT_DIR` | screenshot destination |
| `E2E_LIBRARY` / `E2E_LIBRARY_PATH` | target library (default `Documents` / `/Shared Documents`) |
| `E2E_NAV_WAIT` | per-navigation settle time in ms (default `9000`) |

Credentials are read from the environment only. Never write them to a file in
the repo. The profile directory holds live session cookies and defaults to a
location outside the working tree so it cannot be committed by accident.

## Notes

- **MFA.** The suite drives username/password. If the account enforces MFA, sign
  in once by hand with `E2E_PROFILE_DIR` pointed at a persistent directory; the
  cookies are reused afterwards and the login flow is skipped.
- **Fixtures.** `fixtures.js` uploads two files to the target library on every
  run (overwriting in place). Without at least one item, the column inspector
  reports *"Skipping item data fetch"* and the current-value feature goes
  untested. Point this at a disposable tenant.
- **Inspectors do not close on Escape.** Leaving one open means the next
  `Alt+<key>` toggles it shut, so the panel is in the DOM but hidden and the
  wait for visibility times out. `closePanels()` clicks the real close button
  between scenarios.
- **`chrome.commands` shortcuts are untestable here.** `Alt+N` and `Alt+O` are
  browser-level bindings; synthetic key events dispatched into the page cannot
  reach them. Only page-level listeners (`Ctrl+K`, `Alt+C`, `Alt+I`, `Alt+P`)
  respond to `keyboard.press`. The Lists scenario therefore sends the same
  runtime message the service worker would.

## Seeding a dev tenant

`npm run seed:dev` (`tools/`) builds a site worth inspecting: four groups at
distinct permission levels, five custom lists covering every field type the
inspectors render (Choice, MultiChoice, User, Lookup, Currency, Number,
DateTime, Boolean, Note, URL), a document library with folders and files, and a
`Contracts` list with broken inheritance plus one item carrying unique
permissions. It is idempotent — re-running updates in place.

It cannot create M365 user accounts: that needs Graph with admin consent, and
direct Graph calls from the SharePoint origin return 401. The seeder therefore
assigns whichever real principals the tenant already has to the groups it
creates. Add accounts through Entra first if you want more than one identity.

## Security defaults / MFA

If the tenant has security defaults enabled, Entra eventually forces MFA
enrolment and a plain password login stops working. `ensureSignedIn` checks for
an existing valid session **before** touching the login flow, so a warm profile
keeps working through the intermittent "Action Required" nag.

For unattended runs against an MFA-enforced account, enrol an authenticator and
pass its seed as `SP_TOTP_SECRET`. `totp.js` implements RFC 6238 (30s / 6-digit
/ SHA-1, validated against the spec's test vectors) and `ensureSignedIn` answers
the code challenge automatically:

```
auth: account submitted → password submitted → MFA code submitted → stay-signed-in accepted
```

When enrolling, take the **"Set up a different authentication app"** branch. The
Microsoft Authenticator branch registers a *push* method — it hands out a
numeric activation code, not a TOTP seed, and cannot be answered without a real
device. The branch you want ends on a screen showing `Secret key:` with a 16-
character base32 value.

> `SP_TOTP_SECRET` is a second password with no expiry. Anything holding it can
> complete MFA for that account indefinitely. Use it only for disposable test
> accounts, keep it in the environment rather than a file, and never commit it.
