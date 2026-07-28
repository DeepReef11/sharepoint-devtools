/**
 * SharePoint URL Parser
 * Parses SharePoint URLs and extracts components like site collection, web, list IDs, etc.
 */

import { SharePointUrlComponents } from '../types/sharepoint-context';

/**
 * Regular expressions for SharePoint URL patterns
 */
const SHAREPOINT_PATTERNS = {
  // SharePoint Online domain pattern
  ONLINE_DOMAIN: /^[^.]+\.sharepoint\.com$/i,

  // Site collection pattern - matches /sites/sitename or /teams/teamname
  SITE_COLLECTION: /^\/(sites|teams)\/([^/]+)/i,

  // List/Library patterns
  LIST_FORMS: /\/(Lists|Forms)\/([^/]+)/i,
  LIST_ASPX: /\/AllItems\.aspx|\/DispForm\.aspx|\/EditForm\.aspx|\/NewForm\.aspx/i,

  // List ID in query string
  LIST_ID: /[?&]List=\{?([a-f0-9-]+)\}?/i,

  // Item ID in query string
  ITEM_ID: /[?&]ID=(\d+)/i,

  // View ID in query string
  VIEW_ID: /[?&]View=\{?([a-f0-9-]+)\}?/i,

  // Modern page patterns
  MODERN_PAGE: /\/SitePages\/[^/]+\.aspx/i,

  // Settings pages
  SETTINGS: /\/_layouts\/15\/settings\.aspx/i,
  SITE_CONTENTS: /\/_layouts\/15\/viewlsts\.aspx/i,
};

/**
 * Checks if a URL is a SharePoint URL
 */
export function isSharePointUrl(url: string): boolean {
  try {
    const urlObj = new URL(url);
    return SHAREPOINT_PATTERNS.ONLINE_DOMAIN.test(urlObj.hostname);
  } catch {
    return false;
  }
}

/**
 * Parses query string into a Map
 */
function parseQueryString(queryString: string): Map<string, string> {
  const params = new Map<string, string>();

  if (!queryString || queryString.length === 0) {
    return params;
  }

  // Remove leading '?' if present
  const cleanQuery = queryString.startsWith('?') ? queryString.substring(1) : queryString;

  cleanQuery.split('&').forEach((param) => {
    const [key, value] = param.split('=');
    if (key) {
      params.set(decodeURIComponent(key), value ? decodeURIComponent(value) : '');
    }
  });

  return params;
}

/**
 * Extracts the site collection URL from a SharePoint URL
 * Site collection is typically the root URL or /sites/{sitename} or /teams/{teamname}
 */
function extractSiteCollectionUrl(urlObj: URL): string {
  const pathname = urlObj.pathname;
  const match = pathname.match(SHAREPOINT_PATTERNS.SITE_COLLECTION);

  if (match) {
    // Found /sites/xxx or /teams/xxx pattern
    return `${urlObj.protocol}//${urlObj.hostname}${match[0]}`;
  }

  // Root site collection
  return `${urlObj.protocol}//${urlObj.hostname}`;
}

/**
 * Extracts the web (site/subsite) URL from a SharePoint URL
 * This is more complex as subsites can be nested
 */
function extractWebUrl(urlObj: URL, siteCollectionUrl: string): string {
  const pathname = urlObj.pathname;
  const siteCollectionPath = new URL(siteCollectionUrl).pathname;

  // Get the path after the site collection
  let relativePath = pathname;
  if (siteCollectionPath !== '/') {
    relativePath = pathname.substring(siteCollectionPath.length);
  }

  // Common SharePoint folder patterns that indicate we're in a library/list
  const libraryIndicators = [
    '/Shared Documents',
    '/Documents',
    '/Lists/',
    '/Forms/',
    '/_layouts/',
    '/SitePages/',
    '/SiteAssets/',
    '/Style Library',
  ];

  // Check if path contains library indicators
  for (const indicator of libraryIndicators) {
    const index = relativePath.indexOf(indicator);
    if (index !== -1) {
      // Web URL is everything before the library indicator
      const webPath = relativePath.substring(0, index);
      return siteCollectionUrl + webPath;
    }
  }

  // If no library indicators found, assume the web URL is the site collection
  // This handles root site pages and simple scenarios
  return siteCollectionUrl;
}

/**
 * Extracts the relative path within the web
 */
function extractRelativePath(urlObj: URL, webUrl: string): string {
  const webPath = new URL(webUrl).pathname;
  const fullPath = urlObj.pathname;

  if (fullPath.startsWith(webPath)) {
    const relative = fullPath.substring(webPath.length);
    return relative.startsWith('/') ? relative : '/' + relative;
  }

  return fullPath;
}

/**
 * Extracts tenant ID from SharePoint Online domain
 * For contoso.sharepoint.com, returns "contoso"
 */
export function extractTenantId(hostname: string): string | null {
  if (!SHAREPOINT_PATTERNS.ONLINE_DOMAIN.test(hostname)) {
    return null;
  }

  const parts = hostname.split('.');
  return parts[0] || null;
}

/**
 * Parses a SharePoint URL and extracts all components
 */
export function parseSharePointUrl(url: string): SharePointUrlComponents | null {
  try {
    const urlObj = new URL(url);

    // Verify it's a SharePoint URL
    if (!isSharePointUrl(url)) {
      return null;
    }

    // Extract site collection URL
    const siteCollectionUrl = extractSiteCollectionUrl(urlObj);

    // Extract web URL (may be same as site collection or a subsite)
    const webUrl = extractWebUrl(urlObj, siteCollectionUrl);

    // Extract relative path
    const relativePath = extractRelativePath(urlObj, webUrl);

    // Parse query parameters
    const queryParams = parseQueryString(urlObj.search);

    return {
      fullUrl: url,
      protocol: urlObj.protocol.replace(':', ''),
      domain: urlObj.hostname,
      siteCollectionUrl,
      webUrl,
      relativePath,
      queryParams,
    };
  } catch {
    // Invalid URL
    return null;
  }
}

/**
 * Extracts list ID from URL query parameters or path
 */
export function extractListId(urlComponents: SharePointUrlComponents): string | null {
  // Check query parameters first
  const listParam = urlComponents.queryParams.get('List');
  if (listParam) {
    // Remove curly braces if present
    return listParam.replace(/[{}]/g, '');
  }

  // Check for RootFolder parameter which might contain list name
  // This is a fallback and would need additional API calls to resolve to GUID
  const rootFolder = urlComponents.queryParams.get('RootFolder');
  if (rootFolder) {
    // Extract list name from path like /sites/site/Lists/ListName
    const match = rootFolder.match(/\/Lists\/([^/]+)/i);
    if (match) {
      return match[1]; // Returns list name, not GUID (would need REST API to resolve)
    }
  }

  return null;
}

/**
 * Extracts item ID from URL query parameters
 */
export function extractItemId(urlComponents: SharePointUrlComponents): string | null {
  const idParam = urlComponents.queryParams.get('ID');
  return idParam || null;
}

/**
 * Extracts view ID from URL query parameters
 */
export function extractViewId(urlComponents: SharePointUrlComponents): string | null {
  const viewParam = urlComponents.queryParams.get('View');
  if (viewParam) {
    // Remove curly braces if present
    return viewParam.replace(/[{}]/g, '');
  }
  return null;
}

/**
 * Extracts list name from URL path
 * Note: This returns the URL path name, not the internal name
 */
export function extractListName(urlComponents: SharePointUrlComponents): string | null {
  const path = urlComponents.relativePath;

  // Match /Lists/ListName or /Forms/ListName patterns
  const listMatch = path.match(SHAREPOINT_PATTERNS.LIST_FORMS);
  if (listMatch && listMatch[2]) {
    return listMatch[2];
  }

  // Match library names from common patterns
  const libraryPatterns = [/\/Shared Documents/i, /\/Documents/i, /\/SitePages/i, /\/SiteAssets/i];

  for (const pattern of libraryPatterns) {
    const match = path.match(pattern);
    if (match) {
      return match[0].substring(1); // Remove leading slash
    }
  }

  return null;
}
