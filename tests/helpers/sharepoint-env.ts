/**
 * A fake SharePoint page for jsdom.
 *
 * The extension only activates on a `*.sharepoint.com` hostname and reads its
 * context from `_spPageContextInfo` and the URL, so tests need a document that
 * looks like a real list view. Set the page URL per test file with:
 *
 *   /**
 *    * @jest-environment-options {"url": "https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/AllItems.aspx"}
 *    *\/
 *
 * `installSharePointPage()` then supplies the page context and routes `_api`
 * requests to fixtures, so the real ColumnFetcher parsing code runs unchanged.
 */

import { jest } from '@jest/globals';
import { LIST_ID, WEB_URL } from '../fixtures/sharepoint-api';

export interface ApiRoute {
  /** Substring matched against the request URL. First match wins. */
  match: string;
  /** Parsed JSON body to return, or a status for the failure case. */
  body?: unknown;
  status?: number;
}

export interface SharePointPageOptions {
  webUrl?: string;
  listId?: string;
  routes?: ApiRoute[];
}

/** Requests the stub received, for asserting on endpoints and headers. */
export const apiCalls: { url: string; init?: RequestInit }[] = [];

/**
 * Installs `_spPageContextInfo` and a `fetch` stub for the current test.
 * Returns a teardown function; call it in `afterEach`.
 */
export function installSharePointPage(options: SharePointPageOptions = {}): () => void {
  const { webUrl = WEB_URL, listId = LIST_ID, routes = [] } = options;

  apiCalls.length = 0;

  (window as unknown as Record<string, unknown>)._spPageContextInfo = {
    webAbsoluteUrl: webUrl,
    siteAbsoluteUrl: new URL(webUrl).origin,
    listId: `{${listId}}`,
    isSPO: true,
  };

  const fetchStub = jest.fn(async (input: unknown, init?: RequestInit) => {
    const url = String(input);
    apiCalls.push({ url, init });

    const route = routes.find((r) => url.includes(r.match));
    if (!route) {
      // Surface unmocked endpoints loudly rather than returning undefined.
      throw new Error(`Unmocked SharePoint endpoint: ${url}`);
    }

    const status = route.status ?? 200;
    return {
      ok: status >= 200 && status < 300,
      status,
      statusText: status === 200 ? 'OK' : 'Error',
      json: async () => route.body,
      text: async () => JSON.stringify(route.body),
    };
  });

  (globalThis as unknown as Record<string, unknown>).fetch = fetchStub;

  return () => {
    delete (window as unknown as Record<string, unknown>)._spPageContextInfo;
    delete (globalThis as unknown as Record<string, unknown>).fetch;
    document.body.innerHTML = '';
    document.head.innerHTML = '';
    apiCalls.length = 0;
  };
}

/**
 * Returns every attribute name present anywhere in a subtree, so tests can assert
 * that no event handler was introduced without naming each one individually.
 */
export function attributeNamesIn(root: ParentNode): Set<string> {
  const names = new Set<string>();
  root.querySelectorAll('*').forEach((el) => {
    for (const attr of Array.from(el.attributes)) {
      names.add(attr.name.toLowerCase());
    }
  });
  return names;
}

/** Any `on*` attribute in the subtree — the signature of a successful injection. */
export function eventHandlerAttributes(root: ParentNode): string[] {
  return [...attributeNamesIn(root)].filter((n) => n.startsWith('on'));
}
