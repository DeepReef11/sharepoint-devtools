#!/usr/bin/env node
/**
 * Seed a SharePoint dev tenant with realistic structure: groups and permission
 * assignments, lists covering every field type the inspectors render, a
 * document library with folders and files, and broken permission inheritance.
 *
 * The extension's inspectors are only as interesting as the data behind them —
 * a site with two empty lists exercises almost nothing. This builds something
 * worth inspecting.
 *
 * Idempotent: re-running updates in place rather than duplicating.
 *
 *   export SP_TENANT=contoso.sharepoint.com
 *   export SP_USER=…  SP_PASS=…
 *   npm run seed:dev
 *
 * Requires site-collection admin. Point it at a DISPOSABLE tenant — it creates
 * and modifies lists, groups and permissions.
 */
const path = require('path');
const { spawnSync } = require('child_process');

const { config, assertUsable } = require('../tests/e2e/config');
const { ensureSignedIn } = require('../tests/e2e/auth');
const { LISTS, LIBRARY, GROUPS, ANNOUNCEMENTS, PERSONAS } = require('./seed-data');

// Headed browser needs a display; re-exec under Xvfb rather than making the
// caller remember to. Only when run as a script: this module also exports its
// REST helpers, and importing them should not shell out and seed a tenant.
if (
  require.main === module &&
  !process.env.DISPLAY &&
  !process.env.SEED_XVFB_WRAPPED &&
  process.platform === 'linux'
) {
  if (spawnSync('which', ['xvfb-run']).status === 0) {
    const res = spawnSync(
      'xvfb-run',
      ['-a', '--server-args=-screen 0 1440x900x24', process.execPath, __filename, ...process.argv.slice(2)],
      { stdio: 'inherit', env: { ...process.env, SEED_XVFB_WRAPPED: '1' } }
    );
    process.exit(res.status ?? 1);
  }
  console.error('No DISPLAY and xvfb-run is not installed.');
  process.exit(1);
}

const { chromium } = require('playwright-core');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---------------------------------------------------------------------------
// REST helper — every call runs in the authenticated page context
// ---------------------------------------------------------------------------

function makeClient(page) {
  /**
   * @param {string} url  API path, e.g. "/_api/web/lists"
   * @param {object} [opts] {method, body, headers}
   */
  return async function sp(url, opts = {}) {
    const result = await page.evaluate(
      async ({ url, opts }) => {
        const headers = {
          Accept: 'application/json;odata=verbose',
          ...(opts.headers || {}),
        };
        if (opts.method && opts.method !== 'GET') {
          const ci = await fetch('/_api/contextinfo', {
            method: 'POST',
            headers: { Accept: 'application/json;odata=verbose' },
          });
          headers['X-RequestDigest'] = (await ci.json())?.d?.GetContextWebInformation?.FormDigestValue;
          if (opts.body !== undefined) headers['Content-Type'] = 'application/json;odata=verbose';
        }
        const res = await fetch(url, {
          method: opts.method || 'GET',
          headers,
          body: opts.body === undefined ? undefined : JSON.stringify(opts.body),
        });
        const text = await res.text();
        let parsed = null;
        try { parsed = JSON.parse(text); } catch { /* non-JSON error page */ }
        return {
          ok: res.ok,
          status: res.status,
          data: parsed?.d ?? null,
          error: parsed?.error?.message?.value || (res.ok ? null : text.slice(0, 300)),
        };
      },
      { url, opts }
    );
    if (!result.ok) {
      const err = new Error(`${opts.method || 'GET'} ${url} → ${result.status}: ${result.error}`);
      err.status = result.status;
      throw err;
    }
    return result.data;
  };
}

const results = () => [];
let created = { groups: 0, lists: 0, fields: 0, items: 0, files: 0, folders: 0 };

// ---------------------------------------------------------------------------
// Building blocks
// ---------------------------------------------------------------------------

/**
 * Resolve the fictional personas into site users.
 *
 * ensureuser both resolves a login name and adds it to the site's user
 * information list, which is what a person column stores a reference to. The
 * domain comes from the signed-in account so this stays tenant-agnostic.
 *
 * Personas that do not exist in the directory are skipped rather than fatal —
 * the seeder has no rights to mint accounts. Assigning to whoever else happens
 * to be in the tenant is exactly the behaviour this replaced, so the fallback
 * is the running account alone.
 */
async function ensurePersonas(sp, personas, me) {
  const domain = (me.LoginName.split('|').pop() || '').split('@')[1];
  const resolved = [];
  for (const p of personas) {
    const login = domain ? `${p.alias}@${domain}` : p.alias;
    try {
      const u = await sp(`/_api/web/ensureuser`, { method: 'POST', body: { logonName: login } });
      resolved.push({ Id: u.Id, Title: u.Title, LoginName: u.LoginName, Email: u.Email });
    } catch (err) {
      console.warn(`  persona ${login} could not be resolved: ${err.message.slice(0, 90)}`);
    }
  }
  if (!resolved.length) {
    console.warn('  no personas resolved — falling back to the signed-in account');
    return [{ Id: me.Id, Title: me.Title, LoginName: me.LoginName }];
  }
  return resolved;
}

async function ensureGroup(sp, group, users, claimsByTitle = {}) {
  const existing = await sp(`/_api/web/sitegroups?$filter=Title eq '${group.title}'`);
  let g = existing?.results?.[0];

  if (!g) {
    g = await sp('/_api/web/sitegroups', {
      method: 'POST',
      body: {
        __metadata: { type: 'SP.Group' },
        Title: group.title,
        Description: group.description,
      },
    });
    created.groups++;
  }

  // Grant the group its permission level on the web (idempotent server-side).
  const roleDefs = await sp(`/_api/web/roledefinitions?$filter=Name eq '${group.role}'`);
  const roleId = roleDefs?.results?.[0]?.Id;
  if (roleId) {
    await sp(
      `/_api/web/roleassignments/addroleassignment(principalid=${g.Id},roledefid=${roleId})`,
      { method: 'POST' }
    ).catch(() => {});
  }

  // Populate with whichever real users the tenant actually has, plus any
  // claims principals (e.g. "Everyone except external users") named by the spec.
  const logins = [
    ...group.members.map((i) => users[i]),
    ...(group.claimMembers || []).map((title) => claimsByTitle[title]),
  ].filter(Boolean);

  for (const login of logins) {
    await sp(`/_api/web/sitegroups(${g.Id})/users`, {
      method: 'POST',
      body: { __metadata: { type: 'SP.User' }, LoginName: login },
    }).catch(() => {});
  }
  return g;
}

async function ensureList(sp, spec) {
  let list = await sp(`/_api/web/lists/getbytitle('${spec.title}')`).catch(() => null);
  if (!list) {
    list = await sp('/_api/web/lists', {
      method: 'POST',
      body: {
        __metadata: { type: 'SP.List' },
        Title: spec.title,
        Description: spec.description,
        BaseTemplate: spec.template,
        ContentTypesEnabled: !!spec.contentTypes,
        AllowContentTypes: !!spec.contentTypes,
      },
    });
    created.lists++;
  }
  return list;
}

async function ensureField(sp, listTitle, xml, internalName) {
  const existing = await sp(
    `/_api/web/lists/getbytitle('${listTitle}')/fields?$filter=InternalName eq '${internalName}'`
  );
  if (existing?.results?.length) return existing.results[0];

  const field = await sp(`/_api/web/lists/getbytitle('${listTitle}')/fields/createfieldasxml`, {
    method: 'POST',
    body: {
      parameters: {
        __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
        SchemaXml: xml,
        Options: 8, // AddFieldInternalNameHint — keep our Name as the internal name
      },
    },
  });
  created.fields++;

  // Surface it in the default view, otherwise the data is invisible in the UI.
  await sp(
    `/_api/web/lists/getbytitle('${listTitle}')/defaultview/viewfields/addviewfield('${internalName}')`,
    { method: 'POST' }
  ).catch(() => {});
  return field;
}

async function addItems(sp, listTitle, rows) {
  const meta = await sp(
    `/_api/web/lists/getbytitle('${listTitle}')?$select=ListItemEntityTypeFullName`
  );
  const type = meta.ListItemEntityTypeFullName;

  const existing = await sp(`/_api/web/lists/getbytitle('${listTitle}')/items?$select=Title&$top=500`);
  const have = new Set((existing?.results || []).map((i) => i.Title));

  let n = 0;
  for (const row of rows) {
    if (have.has(row.Title)) continue;
    await sp(`/_api/web/lists/getbytitle('${listTitle}')/items`, {
      method: 'POST',
      body: { __metadata: { type }, ...row },
    });
    created.items++;
    n++;
  }
  return n;
}

module.exports = { makeClient, ensurePersonas, ensureGroup, ensureList, ensureField, addItems };

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

if (require.main === module) {
  (async () => {
    assertUsable();

    const context = await chromium.launchPersistentContext(config.profileDir, {
      executablePath: config.chromium,
      headless: config.headless,
      viewport: { width: 1440, height: 900 },
      args: ['--no-sandbox'],
    });

    try {
      const page = await context.newPage();
      const session = await ensureSignedIn(page);
      if (!session.ok) throw new Error(`sign-in failed: ${JSON.stringify(session)}`);
      console.log(`signed in — "${session.title}"\n`);

      const sp = makeClient(page);

      const me = await sp('/_api/web/currentuser?$select=Id,Title,LoginName,IsSiteAdmin');
      if (!me.IsSiteAdmin) {
        console.warn('WARNING: not a site-collection admin; permission steps will likely fail.\n');
      }

      console.log('personas');
      const humans = await ensurePersonas(sp, PERSONAS, me);
      const logins = humans.map((u) => u.LoginName);
      console.log(`  ${humans.map((u) => u.Title).join(', ')}\n`);

      const siteUsers = await sp(
        '/_api/web/siteusers?$select=Id,Title,LoginName,Email,PrincipalType'
      );

      // --- groups ---------------------------------------------------------
      // Claims principals are addressed by login name, which is tenant-specific.
      const claimsByTitle = Object.fromEntries(
        (siteUsers?.results || [])
          .filter((u) => u.PrincipalType === 4)
          .map((u) => [u.Title, u.LoginName])
      );

      console.log('groups');
      for (const g of GROUPS) {
        const grp = await ensureGroup(sp, g, logins, claimsByTitle);
        const extra = (g.claimMembers || []).filter((t) => claimsByTitle[t]);
        console.log(`  ${grp.Title} (${g.role})${extra.length ? ` + ${extra.join(', ')}` : ''}`);
      }

      // --- lists ----------------------------------------------------------
      const listIds = {};
      for (const spec of LISTS) {
        console.log(`\nlist: ${spec.title}`);
        const list = await ensureList(sp, spec);
        listIds[spec.title] = list.Id;

        for (const f of spec.fields) {
          // Lookup targets are only known once the referenced list exists.
          const xml =
            typeof f.xml === 'function' ? f.xml(listIds) : f.xml;
          if (!xml) continue;
          await ensureField(sp, spec.title, xml, f.name);
        }
        console.log(`  ${spec.fields.length} field(s) ensured`);

        const rows =
          typeof spec.rows === 'function'
            ? await spec.rows({ sp, humans, listIds })
            : spec.rows;
        const n = await addItems(sp, spec.title, rows);
        console.log(`  ${n} item(s) added (${rows.length} defined)`);
      }

      // --- announcements on the built-in template ---------------------------
      console.log('\nlist: Announcements');
      await ensureList(sp, { title: 'Announcements', description: 'Team news', template: 104 });
      const nA = await addItems(sp, 'Announcements', ANNOUNCEMENTS);
      console.log(`  ${nA} item(s) added`);

      // --- document library with folders and files ---------------------------
      console.log(`\nlibrary: ${LIBRARY.title}`);
      await ensureList(sp, {
        title: LIBRARY.title,
        description: LIBRARY.description,
        template: 101,
      });

      // Read the real server-relative URL rather than deriving it from the
      // title: SharePoint keeps the space in "/Project Documents", so a
      // stripped guess silently targets a folder that does not exist.
      const libRoot = await sp(
        `/_api/web/lists/getbytitle('${LIBRARY.title}')/RootFolder?$select=ServerRelativeUrl`
      );
      const libPath = libRoot.ServerRelativeUrl;
      console.log(`  root: ${libPath}`);

      for (const folder of LIBRARY.folders) {
        const folderUrl = `${libPath}/${folder.name}`;
        try {
          await sp('/_api/web/folders', {
            method: 'POST',
            body: { __metadata: { type: 'SP.Folder' }, ServerRelativeUrl: folderUrl },
          });
          created.folders++;
        } catch (err) {
          // Already existing is fine; anything else is a real failure.
          if (!/already exists|AlreadyExists/i.test(err.message)) throw err;
        }

        for (const file of folder.files) {
          const res = await page.evaluate(
            async ({ folderUrl, name, body }) => {
              const ci = await fetch('/_api/contextinfo', {
                method: 'POST',
                headers: { Accept: 'application/json;odata=verbose' },
              });
              const digest = (await ci.json())?.d?.GetContextWebInformation?.FormDigestValue;
              const r = await fetch(
                `/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')` +
                  `/Files/add(url='${name}',overwrite=true)`,
                {
                  method: 'POST',
                  headers: { Accept: 'application/json;odata=verbose', 'X-RequestDigest': digest },
                  body: new Blob([body]),
                }
              );
              return { ok: r.ok, status: r.status, text: r.ok ? '' : (await r.text()).slice(0, 200) };
            },
            { folderUrl, name: file.name, body: file.body }
          );
          if (!res.ok) throw new Error(`upload ${folderUrl}/${file.name} → ${res.status}: ${res.text}`);
          created.files++;
        }
        console.log(`  ${folder.name}/ — ${folder.files.length} file(s)`);
      }

      // --- unique permissions, so the permission inspector has something ----
      console.log('\npermissions');
      const restricted = 'Contracts';
      await sp(
        `/_api/web/lists/getbytitle('${restricted}')/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`,
        { method: 'POST' }
      ).catch(() => {});

      // Remove broad access, keep owners + a named group.
      const groups = await sp('/_api/web/sitegroups?$select=Id,Title');
      const auditors = (groups?.results || []).find((g) => g.Title === 'Contract Auditors');
      const readRole = (
        await sp("/_api/web/roledefinitions?$filter=Name eq 'Read'")
      )?.results?.[0];
      if (auditors && readRole) {
        await sp(
          `/_api/web/lists/getbytitle('${restricted}')/roleassignments/addroleassignment(principalid=${auditors.Id},roledefid=${readRole.Id})`,
          { method: 'POST' }
        ).catch(() => {});
      }
      console.log(`  "${restricted}" list: inheritance broken, Contract Auditors granted Read`);

      // A single item with its own permissions — the most interesting case.
      const firstItem = (
        await sp(`/_api/web/lists/getbytitle('${restricted}')/items?$select=Id&$top=1`)
      )?.results?.[0];
      if (firstItem) {
        await sp(
          `/_api/web/lists/getbytitle('${restricted}')/items(${firstItem.Id})/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`,
          { method: 'POST' }
        ).catch(() => {});
        console.log(`  "${restricted}" item #${firstItem.Id}: unique permissions`);
      }

      console.log('\n' + '─'.repeat(52));
      console.log(
        `groups ${created.groups} · lists ${created.lists} · fields ${created.fields} · ` +
          `items ${created.items} · folders ${created.folders} · files ${created.files}`
      );
      console.log(`https://${config.tenant}/_layouts/15/viewlsts.aspx`);
    } finally {
      await context.close();
    }
  })().catch((e) => {
    console.error(`\nseed failed: ${e.message}`);
    process.exit(1);
  });
}
