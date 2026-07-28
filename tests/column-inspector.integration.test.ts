/**
 * @jest-environment-options {"url": "https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/DispForm.aspx?ID=42"}
 */

/**
 * Column Inspector integration tests.
 *
 * These drive the real component against mocked REST responses: context detection,
 * ColumnFetcher parsing, and DOM rendering all run unmodified. No tenant required.
 */

import { describe, it, expect, beforeEach, afterEach } from '@jest/globals';
import { ColumnInspector } from '../src/ui/column-inspector';
import {
  installSharePointPage,
  eventHandlerAttributes,
  attributeNamesIn,
  type ApiRoute,
} from './helpers/sharepoint-env';
import {
  HOSTILE,
  LIST_SCHEMA_RESPONSE,
  LIST_ITEM_RESPONSE,
  SAFE_URL_ITEM_RESPONSE,
  WEB_URL,
} from './fixtures/sharepoint-api';

const SCHEMA_ROUTE: ApiRoute = { match: '$expand=Fields', body: LIST_SCHEMA_RESPONSE };
const ITEM_ROUTE: ApiRoute = { match: '/items(', body: LIST_ITEM_RESPONSE };

let teardown: () => void;

async function openInspector(routes: ApiRoute[] = [SCHEMA_ROUTE, ITEM_ROUTE]) {
  teardown = installSharePointPage({ routes });
  const inspector = new ColumnInspector();
  await inspector.show();
  return inspector;
}

function inspectorRoot(): HTMLElement {
  const el = document.getElementById('sp-column-inspector');
  if (!el) throw new Error('inspector not rendered');
  return el;
}

afterEach(() => teardown?.());

describe('ColumnInspector rendering', () => {
  it('renders the list schema fetched from the REST API', async () => {
    await openInspector();
    const text = inspectorRoot().textContent ?? '';

    expect(text).toContain('Project Tasks');
    expect(text).toContain('ProjectTasksList');
  });

  it('renders one card per user-visible column', async () => {
    await openInspector();
    const cards = inspectorRoot().querySelectorAll('.sp-column-card');

    // 17 fields in the fixture, minus 2 filtered system fields, minus the
    // hidden column, which is excluded until "Show hidden" is enabled.
    expect(cards.length).toBe(14);
    expect(inspectorRoot().textContent).not.toContain('Hidden Field');
  });

  it('omits system fields that ColumnFetcher filters out', async () => {
    await openInspector();
    const text = inspectorRoot().textContent ?? '';

    expect(text).not.toContain('owshiddenversion');
  });

  it('surfaces calculated column formulas and lookup targets', async () => {
    await openInspector();
    const text = inspectorRoot().textContent ?? '';

    expect(text).toContain('=[Budget]*1.15');
    expect(text).toContain('Department');
  });

  it('requests the expected endpoint with SharePoint credentials', async () => {
    await openInspector();
    const { apiCalls } = await import('./helpers/sharepoint-env');
    const schemaCall = apiCalls.find((c) => c.url.includes('$expand=Fields'));

    expect(schemaCall).toBeDefined();
    expect(schemaCall!.url).toContain(`${WEB_URL}/_api/web/lists(guid'`);
    expect(schemaCall!.init?.credentials).toBe('include');
  });
});

describe('ColumnInspector against hostile SharePoint data', () => {
  // Every string below is something any site contributor can store in a list.
  // The inspector renders it into the page, so it must never become live markup.

  it('introduces no event-handler attributes anywhere in the panel', async () => {
    await openInspector();

    expect(eventHandlerAttributes(inspectorRoot())).toEqual([]);
  });

  it('does not create elements from markup in a column title', async () => {
    await openInspector();
    const root = inspectorRoot();

    expect(root.querySelector('script')).toBeNull();
    expect(root.querySelector('img')).toBeNull();
  });

  it('renders an attribute-breakout payload as text, not as an attribute', async () => {
    await openInspector();
    const root = inspectorRoot();

    // The payload should be visible to the user as literal text...
    expect(root.textContent).toContain(HOSTILE.attributeBreakout);
    // ...and must not have become a real attribute. Checking the parsed DOM
    // rather than the innerHTML string: correctly escaped markup still contains
    // the substring "onmouseover=" harmlessly, so a text search proves nothing.
    expect(attributeNamesIn(root).has('onmouseover')).toBe(false);
  });

  it('does not render a javascript: URL column value as a link', async () => {
    await openInspector();
    const hrefs = Array.from(inspectorRoot().querySelectorAll('a')).map((a) =>
      a.getAttribute('href')
    );

    expect(hrefs.some((h) => h?.toLowerCase().startsWith('javascript:'))).toBe(false);
    expect(hrefs.some((h) => h?.toLowerCase().startsWith('data:'))).toBe(false);
  });

  it('still renders a legitimate URL column value as a working link', async () => {
    await openInspector([SCHEMA_ROUTE, { match: '/items(', body: SAFE_URL_ITEM_RESPONSE }]);
    const hrefs = Array.from(inspectorRoot().querySelectorAll('a')).map((a) =>
      a.getAttribute('href')
    );

    expect(hrefs).toContain('https://contoso.sharepoint.com/sites/marketing/docs');
  });

  it('keeps hostile content-type names inert', async () => {
    await openInspector();
    const root = inspectorRoot();

    expect(root.querySelector('img')).toBeNull();
    expect(attributeNamesIn(root).has('onerror')).toBe(false);
  });
});

describe('ColumnInspector error handling', () => {
  it('shows an error instead of throwing when the API rejects the request', async () => {
    teardown = installSharePointPage({
      routes: [{ match: '$expand=Fields', status: 403, body: null }],
    });

    const inspector = new ColumnInspector();
    await expect(inspector.show()).resolves.not.toThrow();

    const errorEl = document.getElementById('sp-inspector-error');
    expect(errorEl?.style.display).not.toBe('none');
  });
});
