/**
 * Seed the target library with documents.
 *
 * Without at least one item, the column inspector's item-data path never runs
 * ("Skipping item data fetch - no item ID/path"), so the current-value feature
 * would silently go untested.
 */
const { config } = require('./config');

const FIXTURES = [
  {
    name: 'quicknav-e2e-alpha.txt',
    body: 'QuickNav e2e fixture A.\nExercises the column inspector item-value path.\n',
  },
  {
    name: 'quicknav-e2e-beta.csv',
    body: 'id,name,qty\n1,widget,10\n2,gasket,4\n3,flange,7\n',
  },
];

/**
 * Upload the fixtures (overwriting any previous copy) and return the resulting
 * list items so scenarios can address them by numeric ID.
 */
async function seedLibrary(page, { log = console.log } = {}) {
  await page.goto(`https://${config.tenant}/`, {
    waitUntil: 'domcontentloaded',
    timeout: 60000,
  });
  await new Promise((r) => setTimeout(r, 5000));

  const uploaded = await page.evaluate(
    async ({ files, libraryPath }) => {
      const info = await fetch('/_api/contextinfo', {
        method: 'POST',
        headers: { Accept: 'application/json;odata=verbose' },
      });
      const digest = (await info.json())?.d?.GetContextWebInformation?.FormDigestValue;
      if (!digest) return { error: 'could not obtain form digest' };

      const results = [];
      for (const file of files) {
        const res = await fetch(
          `/_api/web/GetFolderByServerRelativeUrl('${libraryPath}')` +
            `/Files/add(url='${file.name}',overwrite=true)`,
          {
            method: 'POST',
            headers: { Accept: 'application/json;odata=verbose', 'X-RequestDigest': digest },
            body: new Blob([file.body]),
          }
        );
        results.push({ name: file.name, status: res.status });
      }
      return { results };
    },
    { files: FIXTURES, libraryPath: config.libraryPath }
  );

  if (uploaded.error) throw new Error(`Fixture upload failed: ${uploaded.error}`);
  const failed = uploaded.results.filter((r) => r.status >= 300);
  if (failed.length) {
    throw new Error(`Fixture upload rejected: ${JSON.stringify(failed)}`);
  }
  log(`  fixtures: ${uploaded.results.length} uploaded`);

  const items = await page.evaluate(async (library) => {
    const res = await fetch(
      `/_api/web/lists/getbytitle('${library}')/items?$select=Id,FileLeafRef&$top=50`,
      { headers: { Accept: 'application/json;odata=verbose' } }
    );
    const body = await res.json();
    return (body?.d?.results || []).map((i) => ({ id: i.Id, name: i.FileLeafRef }));
  }, config.library);

  if (!items.length) throw new Error('Library still reports no items after seeding');
  log(`  fixtures: library has ${items.length} item(s)`);
  return items;
}

module.exports = { seedLibrary, FIXTURES };
