/**
 * SharePoint Context Detection Module
 * Extracts and manages SharePoint context information from the current page
 */

/* eslint-disable @typescript-eslint/no-explicit-any */

export interface SharePointContext {
  isSharePoint: boolean;
  siteUrl?: string;
  webUrl?: string;
  listId?: string;
  listUrl?: string;
  pageType?: 'site' | 'list' | 'library' | 'page' | 'unknown';
  isModern?: boolean;
}

/**
 * Detects if the current page is a SharePoint page
 */
export function isSharePointPage(): boolean {
  const url = window.location.href;
  return url.includes('.sharepoint.com') || url.includes('.sharepoint-df.com');
}

/**
 * Extracts the site collection URL from the current page
 * Example: https://tenant.sharepoint.com/sites/sitename
 */
export function extractSiteUrl(): string | undefined {
  const url = window.location.href;
  const match = url.match(/(https?:\/\/[^/]+\/sites\/[^/]+)/);
  if (match) {
    return match[1];
  }

  // Root site collection
  const rootMatch = url.match(/(https?:\/\/[^/]+)/);
  if (rootMatch && isSharePointPage()) {
    return rootMatch[1];
  }

  return undefined;
}

/**
 * Extracts the web URL from the current page
 * For subsites, this will be different from the site URL
 */
export function extractWebUrl(): string | undefined {
  // Try to get from _spPageContextInfo if available
  if (typeof (window as any)._spPageContextInfo !== 'undefined') {
    const pageContextWebUrl = (window as any)._spPageContextInfo.webAbsoluteUrl;
    // Verify it's a valid URL before trusting it
    if (pageContextWebUrl && pageContextWebUrl.includes('/sites/')) {
      return pageContextWebUrl;
    }
  }

  // Fallback to URL parsing with path stripping
  try {
    const urlObj = new URL(window.location.href);
    const pathname = urlObj.pathname;

    // Special handling for document libraries accessed via /LibraryName/Forms/
    // Patterns:
    // - /sites/SITE/LibraryName/Forms/ -> web is /sites/SITE
    // - /sites/SITE/subsite/LibraryName/Forms/ -> web is /sites/SITE/subsite
    // - /sites/SITE/subsite1/subsite2/LibraryName/Forms/ -> web is /sites/SITE/subsite1/subsite2
    // The segment immediately before /Forms/ is the library name, everything before that is the web
    // Match everything up to but NOT including the last segment before /Forms/
    const docLibMatch = pathname.match(/^(.+)\/[^/]+\/Forms\//i);
    if (docLibMatch && docLibMatch[1].includes('/sites/')) {
      return `${urlObj.protocol}//${urlObj.hostname}${docLibMatch[1]}`;
    }

    // Web URL can be a subsite: /sites/sitename/subweb
    // We need to capture everything before known SharePoint paths
    // Known paths: _layouts, Lists, SitePages, Shared Documents, Documents, Forms, etc.
    const knownPaths =
      /\/(Lists|_layouts|SitePages|Shared%20Documents|Shared Documents|Documents|Forms|_api|_catalogs|_vti_bin)($|\/)/i;

    let webPath = pathname;
    const knownPathMatch = pathname.match(knownPaths);
    if (knownPathMatch && knownPathMatch.index !== undefined) {
      // Extract everything before the known path
      webPath = pathname.substring(0, knownPathMatch.index);
    }

    // Match /sites/sitename or /sites/sitename/subweb/subsubweb
    const webMatch = webPath.match(/^(\/sites\/[^/]+(?:\/[^/]+)*)/);
    if (webMatch) {
      return `${urlObj.protocol}//${urlObj.hostname}${webMatch[1]}`;
    }

    // If no subsite, web URL is same as site URL
    return extractSiteUrl();
  } catch (error) {
    console.error('Error extracting web URL:', error);
    return extractSiteUrl();
  }
}

/**
 * Extracts the list name from classic SharePoint list URLs
 * Example: /sites/site/Lists/ListName/view.aspx -> ListName
 */
function extractListNameFromUrl(): string | undefined {
  const pathname = window.location.pathname;

  // Pattern for /Lists/ListName/ or /Lists/ListName/Forms/
  const listsMatch = pathname.match(/\/Lists\/([^/]+)\//i);
  if (listsMatch) {
    return decodeURIComponent(listsMatch[1]);
  }

  // Pattern for document libraries (often don't have /Lists/)
  // Example: /sites/site/LibraryName/Forms/AllItems.aspx
  const formsMatch = pathname.match(
    /\/([^/]+)\/Forms\/(AllItems|DispForm|EditForm|NewForm)\.aspx/i
  );
  if (formsMatch) {
    return decodeURIComponent(formsMatch[1]);
  }

  return undefined;
}

// Cache for list ID lookups to avoid repeated API calls
const listIdCache = new Map<string, string>();

// Global cached list ID for current page (set after async fetch)
let cachedCurrentListId: string | undefined = undefined;

/**
 * Set the cached list ID for the current page
 * This is used after async fetch to make it available synchronously
 */
export function setCachedListId(listId: string | undefined): void {
  cachedCurrentListId = listId;
}

/**
 * Fetches the list GUID from SharePoint REST API using the list title
 */
async function fetchListIdByTitle(webUrl: string, listTitle: string): Promise<string | undefined> {
  const cacheKey = `${webUrl}|${listTitle}`;

  // Check cache first
  if (listIdCache.has(cacheKey)) {
    return listIdCache.get(cacheKey);
  }

  // Try method 1: getByTitle
  try {
    const apiUrl = `${webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')?$select=Id`;

    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        Accept: 'application/json;odata=verbose',
      },
      credentials: 'same-origin',
    });

    if (response.ok) {
      const data = await response.json();
      const listId = data.d?.Id;

      if (listId) {
        const extracted = listId.replace(/[{}]/g, '').toLowerCase();
        listIdCache.set(cacheKey, extracted);
        return extracted;
      }
    }
  } catch {
    // Silently fall through to next method
  }

  // Try method 2: Filter by RootFolder ServerRelativeUrl
  try {
    // Construct the expected server relative URL for the library or list
    const pathname = window.location.pathname;
    let listPath: string | undefined;

    // Pattern 1: Document library - /sites/SITE/LibraryName/Forms/View.aspx
    const libraryPathMatch = pathname.match(/^(.+)\/Forms\//i);
    if (libraryPathMatch) {
      listPath = libraryPathMatch[1];
    }

    // Pattern 2: List - /sites/SITE/Lists/ListName/View.aspx
    // Extract everything before the view .aspx file
    const listsPathMatch = pathname.match(/^(.+\/Lists\/[^/]+)\//i);
    if (!listPath && listsPathMatch) {
      listPath = listsPathMatch[1];
    }

    if (listPath) {
      const filterUrl = `${webUrl}/_api/web/lists?$filter=RootFolder/ServerRelativeUrl eq '${listPath}'&$select=Id,Title`;

      const response = await fetch(filterUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json;odata=verbose',
        },
        credentials: 'same-origin',
      });

      if (response.ok) {
        const data = await response.json();
        const lists = data.d?.results;

        if (lists && lists.length > 0) {
          const listId = lists[0].Id;
          const extracted = listId.replace(/[{}]/g, '').toLowerCase();
          listIdCache.set(cacheKey, extracted);
          return extracted;
        }
      }
    }
  } catch {
    // Silently fall through
  }

  return undefined;
}

/**
 * Extracts the list ID from the current page
 * Looks for list GUID in URL or page context
 */
export function extractListId(): string | undefined {
  // First check if we have a cached ID from async fetch
  if (cachedCurrentListId) {
    return cachedCurrentListId;
  }

  // Try to get from _spPageContextInfo first (most reliable)
  if (typeof (window as any)._spPageContextInfo !== 'undefined') {
    const pageContext = (window as any)._spPageContextInfo;

    if (pageContext.listId) {
      const extracted = pageContext.listId.replace(/[{}]/g, '').toLowerCase();
      return extracted;
    }

    // Try pageListId as fallback
    if (pageContext.pageListId) {
      const extracted = pageContext.pageListId.replace(/[{}]/g, '').toLowerCase();
      return extracted;
    }
  }

  // Try to extract from DOM
  // Modern SharePoint often has list ID in data attributes
  const listElements = [
    document.querySelector('[data-list-id]'),
    document.querySelector('[data-listid]'),
    document.querySelector('#ListWebPart'),
    document.querySelector('[id^="WebPart"]'),
  ];

  for (const elem of listElements) {
    if (elem) {
      const listIdAttr =
        elem.getAttribute('data-list-id') ||
        elem.getAttribute('data-listid') ||
        elem.getAttribute('listid');

      if (listIdAttr) {
        return listIdAttr.replace(/[{}]/g, '').toLowerCase();
      }
    }
  }

  // Check hidden form fields (classic SharePoint)
  const hiddenListId = document.querySelector<HTMLInputElement>(
    'input[name="ListId"], input[id*="ListId"]'
  );
  if (hiddenListId && hiddenListId.value) {
    return hiddenListId.value.replace(/[{}]/g, '').toLowerCase();
  }

  // Check URL for list parameter
  const urlParams = new URLSearchParams(window.location.search);
  const listParam = urlParams.get('List');
  if (listParam) {
    return listParam.replace(/[{}%]/g, '').toLowerCase();
  }

  // Look for list ID in URL path (for modern experiences)
  const guidMatch = window.location.pathname.match(
    /[{]?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}[}]?/
  );
  if (guidMatch) {
    return guidMatch[0].replace(/[{}]/g, '').toLowerCase();
  }

  // Check full URL for GUID (sometimes in hash/query)
  const fullGuidMatch = window.location.href.match(
    /[{]?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}[}]?/
  );
  if (fullGuidMatch) {
    return fullGuidMatch[0].replace(/[{}]/g, '').toLowerCase();
  }

  return undefined;
}

/**
 * Asynchronously extracts the list ID, including fetching from API if needed
 * Use this when you need the list ID but can handle async/await
 */
export async function extractListIdAsync(): Promise<string | undefined> {
  // First try synchronous extraction
  const syncId = extractListId();
  if (syncId) {
    return syncId;
  }

  // If no GUID found, try to get it via API using list name
  const listName = extractListNameFromUrl();
  if (!listName) {
    return undefined;
  }

  // Get web URL for API call
  const siteUrl = extractSiteUrl();
  if (!siteUrl) {
    return undefined;
  }

  return await fetchListIdByTitle(siteUrl, listName);
}

/**
 * Extracts the list URL from the current page
 */
export function extractListUrl(): string | undefined {
  if (typeof (window as any)._spPageContextInfo !== 'undefined') {
    const pageContext = (window as any)._spPageContextInfo;
    if (pageContext.listUrl) {
      const webUrl = pageContext.webAbsoluteUrl;
      return `${webUrl}${pageContext.listUrl}`;
    }
  }

  return undefined;
}

/**
 * Detects the type of SharePoint page
 */
export function detectPageType(): 'site' | 'list' | 'library' | 'page' | 'unknown' {
  const url = window.location.href.toLowerCase();
  const path = window.location.pathname.toLowerCase();

  // Check _spPageContextInfo first (most reliable)
  if (typeof (window as any)._spPageContextInfo !== 'undefined') {
    const pageContext = (window as any)._spPageContextInfo;
    if (pageContext.listId) {
      // Check if it's a document library
      if (
        pageContext.listBaseType === 1 ||
        path.includes('/documents/') ||
        path.includes('/shared%20documents/')
      ) {
        return 'library';
      }
      return 'list';
    }
  }

  // Check for list/library views in URL
  if (
    path.includes('/lists/') ||
    path.includes('/allitems.aspx') ||
    url.includes('viewid=') ||
    url.includes('list=')
  ) {
    // Check if it's a document library
    if (
      path.includes('/documents/') ||
      path.includes('/shared%20documents/') ||
      path.includes('/forms/')
    ) {
      return 'library';
    }
    return 'list';
  }

  // Check for pages
  if (path.includes('/sitepages/') || path.includes('/pages/')) {
    return 'page';
  }

  // Check for site home
  if (
    path.endsWith('/') ||
    path.includes('/_layouts/15/start.aspx') ||
    path.includes('/sitehome.aspx')
  ) {
    return 'site';
  }

  return 'unknown';
}

/**
 * Detects if the page is using modern or classic SharePoint UI
 */
export function detectModernUI(): boolean {
  // Check for modern UI indicators
  const modernIndicators = [
    document.querySelector('[data-sp-feature-tag="modern"]'),
    document.querySelector('.ms-SPLegacyFabricBlock'),
    document.querySelector('[data-automationid="CanvasZone"]'),
  ];

  return modernIndicators.some((indicator) => indicator !== null);
}

/**
 * Gets the complete SharePoint context for the current page
 */
export function getSharePointContext(): SharePointContext {
  if (!isSharePointPage()) {
    return { isSharePoint: false };
  }

  return {
    isSharePoint: true,
    siteUrl: extractSiteUrl(),
    webUrl: extractWebUrl(),
    listId: extractListId(),
    listUrl: extractListUrl(),
    pageType: detectPageType(),
    isModern: detectModernUI(),
  };
}
