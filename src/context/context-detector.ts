/**
 * SharePoint Context Detector
 * Main module for detecting SharePoint context, version, and page type
 */

import {
  SharePointContext,
  SharePointVersion,
  SharePointPageType,
  PageContext,
  ListContext,
  ContextDetectionResult,
  SharePointUrlComponents,
} from '../types/sharepoint-context';
import {
  parseSharePointUrl,
  isSharePointUrl,
  extractTenantId,
  extractListId,
  extractItemId,
  extractViewId,
  extractListName,
} from './url-parser';

/**
 * Maps detailed PageContext to simplified SharePointPageType
 */
function mapPageContextToPageType(pageContext: PageContext): SharePointPageType {
  switch (pageContext) {
    case PageContext.DocumentLibrary:
      return SharePointPageType.Library;
    case PageContext.List:
    case PageContext.ListItem:
      return SharePointPageType.List;
    case PageContext.SiteHome:
    case PageContext.SiteSettings:
    case PageContext.SiteContents:
      return SharePointPageType.Site;
    case PageContext.ModernPage:
    case PageContext.WikiPage:
    case PageContext.WebPartPage:
      return SharePointPageType.Page;
    default:
      return SharePointPageType.Unknown;
  }
}

/**
 * Detects SharePoint version (Modern vs Classic) based on URL patterns and DOM
 * This detection can be enhanced when running in browser context with DOM access
 */
export function detectSharePointVersion(url: string, document?: Document): SharePointVersion {
  // Modern UI indicators in URL
  const modernIndicators = ['/SitePages/', '/_layouts/15/sharepoint.aspx', '/_layouts/15/me.aspx'];

  // Classic UI indicators in URL
  const classicIndicators = [
    '/Forms/AllItems.aspx',
    '/Forms/DispForm.aspx',
    '/Forms/EditForm.aspx',
    '/Forms/NewForm.aspx',
    '/_layouts/15/start.aspx',
    '/_layouts/15/settings.aspx',
  ];

  // Check URL patterns first
  for (const indicator of modernIndicators) {
    if (url.includes(indicator)) {
      return SharePointVersion.Modern;
    }
  }

  for (const indicator of classicIndicators) {
    if (url.includes(indicator)) {
      return SharePointVersion.Classic;
    }
  }

  // If document is provided, check DOM for modern UI indicators
  if (document) {
    // Modern SharePoint uses suite nav
    const suiteNav = document.querySelector('#suiteNav, [data-automationid="suiteNav"]');
    if (suiteNav) {
      return SharePointVersion.Modern;
    }

    // Modern SharePoint uses the modern command bar
    const modernCommandBar = document.querySelector('[data-automationid="commandBar"]');
    if (modernCommandBar) {
      return SharePointVersion.Modern;
    }

    // Classic SharePoint uses ribbon
    const ribbon = document.querySelector('#RibbonContainer, .ms-cui-ribbon');
    if (ribbon) {
      return SharePointVersion.Classic;
    }

    // Check for modern page chrome
    const modernChrome = document.querySelector('.ms-compositeHeader, [data-sp-feature-tag]');
    if (modernChrome) {
      return SharePointVersion.Modern;
    }
  }

  // Default to Modern for SharePoint Online (most common now)
  // If running on-premises, this would need additional detection
  return SharePointVersion.Modern;
}

/**
 * Detects the page context type based on URL and DOM
 */
export function detectPageContext(
  url: string,
  urlComponents: SharePointUrlComponents | null,
  document?: Document
): PageContext {
  const path = urlComponents?.relativePath || '';
  const hasListId = urlComponents && extractListId(urlComponents);
  const hasItemId = urlComponents && extractItemId(urlComponents);

  // Check for specific page patterns
  if (url.includes('/_layouts/15/settings.aspx')) {
    return PageContext.SiteSettings;
  }

  if (url.includes('/_layouts/15/viewlsts.aspx')) {
    return PageContext.SiteContents;
  }

  // Modern page
  if (path.match(/\/SitePages\/.*\.aspx$/i)) {
    return PageContext.ModernPage;
  }

  // List item form (display, edit, new)
  if (
    hasItemId &&
    (url.includes('DispForm.aspx') || url.includes('EditForm.aspx') || url.includes('NewForm.aspx'))
  ) {
    return PageContext.ListItem;
  }

  // Document library
  if (
    path.includes('/Shared Documents') ||
    path.includes('/Documents/') ||
    (path.includes('/Forms/') && url.includes('AllItems.aspx'))
  ) {
    return PageContext.DocumentLibrary;
  }

  // List
  if (path.includes('/Lists/') || hasListId) {
    return PageContext.List;
  }

  // Wiki page
  if (path.match(/\/.*wiki.*\.aspx$/i)) {
    return PageContext.WikiPage;
  }

  // Web part page
  if (path.match(/\/.*\.aspx$/i) && !path.includes('/SitePages/')) {
    return PageContext.WebPartPage;
  }

  // Check if we're on root or home
  if (path === '/' || path === '' || path === '/SitePages/Home.aspx') {
    return PageContext.SiteHome;
  }

  // If DOM is available, do additional detection
  if (document) {
    // Check for list view web part
    const listView = document.querySelector('[data-automationid="listView"], .ms-listviewtable');
    if (listView) {
      // Determine if it's a library or list
      const isLibrary = document.querySelector('[data-automationid="libraryCommandBar"]');
      return isLibrary ? PageContext.DocumentLibrary : PageContext.List;
    }
  }

  return PageContext.Unknown;
}

/**
 * Builds list context information from URL components
 */
export function buildListContext(
  urlComponents: SharePointUrlComponents | null
): ListContext | null {
  if (!urlComponents) {
    return null;
  }

  const listId = extractListId(urlComponents);
  const itemId = extractItemId(urlComponents);
  const viewId = extractViewId(urlComponents);
  const listName = extractListName(urlComponents);

  // Only return list context if we have at least one piece of list information
  if (!listId && !itemId && !viewId && !listName) {
    return null;
  }

  return {
    listId,
    listName,
    listTitle: null, // Would need REST API call to get display name
    itemId,
    viewId,
  };
}

/**
 * Detects complete SharePoint context from current page
 *
 * @param url - The current page URL
 * @param document - Optional document object for DOM-based detection
 * @returns Complete SharePoint context information
 */
export function detectSharePointContext(url: string, document?: Document): ContextDetectionResult {
  try {
    // Check if it's a SharePoint URL
    const isSharePoint = isSharePointUrl(url);

    if (!isSharePoint) {
      return {
        success: false,
        context: null,
        error: 'Not a SharePoint URL',
      };
    }

    // Parse URL components
    const urlComponents = parseSharePointUrl(url);

    if (!urlComponents) {
      return {
        success: false,
        context: null,
        error: 'Failed to parse SharePoint URL',
      };
    }

    // Detect version
    const version = detectSharePointVersion(url, document);

    // Detect page context
    const pageContext = detectPageContext(url, urlComponents, document);

    // Build list context
    const listContext = buildListContext(urlComponents);

    // Extract tenant ID
    const tenantId = extractTenantId(urlComponents.domain);

    // Map PageContext to SharePointPageType for compatibility
    const pageType = mapPageContextToPageType(pageContext);

    // Build complete context
    const context: SharePointContext = {
      isSharePoint: true,
      version,
      pageContext,
      pageType,
      tenantUrl: `${urlComponents.protocol}://${urlComponents.domain}`,
      tenantId,
      siteUrl: urlComponents.siteCollectionUrl,
      webUrl: urlComponents.webUrl,
      listId: listContext?.listId || null,
      listTitle: listContext?.listTitle || null,
      urlComponents,
      listContext,
      isSharePointPage: true,
      currentUrl: url,
      detectedAt: new Date(),
    };

    return {
      success: true,
      context,
      error: null,
    };
  } catch (error) {
    return {
      success: false,
      context: null,
      error: error instanceof Error ? error.message : 'Unknown error occurred',
    };
  }
}

/**
 * Convenience function to detect context from window.location
 * This would be used in the browser extension's content script
 */
export function detectCurrentContext(document?: Document): ContextDetectionResult {
  if (typeof window === 'undefined' || !window.location) {
    return {
      success: false,
      context: null,
      error: 'Not running in browser context',
    };
  }

  return detectSharePointContext(window.location.href, document || window.document);
}

/**
 * Validates if a context object has minimum required information
 */
export function isValidContext(context: SharePointContext | null): boolean {
  if (!context) {
    return false;
  }

  return (
    context.isSharePoint &&
    context.urlComponents !== null &&
    context.urlComponents.webUrl !== null &&
    context.urlComponents.siteCollectionUrl !== null
  );
}
