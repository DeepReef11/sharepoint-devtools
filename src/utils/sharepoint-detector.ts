/**
 * SharePoint page detection and context extraction utilities
 */

import { isSharePointHostname } from './sharepoint-clouds';
import {
  SharePointContext,
  SharePointVersion,
  SharePointPageType,
  SHAREPOINT_URL_PATTERNS,
} from '../types/sharepoint';

/**
 * Detects if the current page is a SharePoint page
 */
export function isSharePointPage(url: string = window.location.href): boolean {
  return (
    SHAREPOINT_URL_PATTERNS.sharepoint.test(url) || SHAREPOINT_URL_PATTERNS.sharepointDf.test(url)
  );
}

/**
 * Detects SharePoint version (Modern vs Classic)
 */
export function detectSharePointVersion(url: string = window.location.href): SharePointVersion {
  // Check for modern UI indicators
  const hasModernSuiteNav = document.querySelector('.od-SuiteNav');
  const hasModernCommandBar = document.querySelector('.ms-CommandBar');
  const hasReactRoot = document.querySelector('[data-sp-feature-tag]');

  // Check URL patterns
  const isModernUrl = url.includes('/_layouts/15/') && !url.includes('/_layouts/15/start.aspx');

  if (hasModernSuiteNav || hasModernCommandBar || hasReactRoot || isModernUrl) {
    return SharePointVersion.MODERN;
  }

  // Check for classic UI indicators
  const hasClassicRibbon = document.querySelector('#RibbonContainer');
  const hasClassicSuiteNav = document.querySelector('#suiteBarLeft');

  if (hasClassicRibbon || hasClassicSuiteNav) {
    return SharePointVersion.CLASSIC;
  }

  return SharePointVersion.UNKNOWN;
}

/**
 * Extracts the site URL from the current page
 */
export function extractSiteUrl(url: string = window.location.href): string | undefined {
  try {
    const urlObj = new URL(url);
    const hostname = urlObj.hostname;
    const pathname = urlObj.pathname;

    // SharePoint site URL is typically: https://tenant.sharepoint.com/sites/sitename
    const siteMatch = pathname.match(/^(\/sites\/[^/]+)/);
    if (siteMatch) {
      return `${urlObj.protocol}//${hostname}${siteMatch[1]}`;
    }

    // Root site collection
    if (isSharePointHostname(hostname)) {
      return `${urlObj.protocol}//${hostname}`;
    }

    return undefined;
  } catch (error) {
    console.error('Error extracting site URL:', error);
    return undefined;
  }
}

/**
 * Extracts the web URL from the current page
 */
export function extractWebUrl(url: string = window.location.href): string | undefined {
  try {
    const urlObj = new URL(url);
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
    // The ($|\/) at the end handles cases with or without trailing slash
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
    return extractSiteUrl(url);
  } catch (error) {
    console.error('Error extracting web URL:', error);
    return undefined;
  }
}

/**
 * Extracts the list ID from the current page
 */
export function extractListId(url: string = window.location.href): string | undefined {
  try {
    // List ID is typically in the URL as a GUID parameter
    const urlObj = new URL(url);
    const params = new URLSearchParams(urlObj.search);

    // Common parameter names for list ID
    const listIdParam = params.get('List') || params.get('ListId') || params.get('list');

    if (listIdParam) {
      // Remove any curly braces and convert to uppercase
      return listIdParam.replace(/[{}]/g, '').toUpperCase();
    }

    // Try to extract from path for modern pages
    const pathMatch = url.match(/\/Lists\/([^/]+)/);
    if (pathMatch) {
      return pathMatch[1];
    }

    return undefined;
  } catch (error) {
    console.error('Error extracting list ID:', error);
    return undefined;
  }
}

/**
 * Detects the type of SharePoint page
 */
export function detectPageType(url: string = window.location.href): SharePointPageType {
  const urlObj = new URL(url);

  // Check for admin sites (domain contains -admin)
  if (urlObj.hostname.includes('-admin.sharepoint.com')) {
    return SharePointPageType.ADMIN;
  }

  // Check for site settings first (most specific)
  if (SHAREPOINT_URL_PATTERNS.siteSettings.test(url)) {
    return SharePointPageType.SITE_SETTINGS;
  }

  // Check for list forms before list views (DispForm, EditForm, NewForm)
  if (
    url.includes('DispForm.aspx') ||
    url.includes('EditForm.aspx') ||
    url.includes('NewForm.aspx')
  ) {
    return SharePointPageType.LIST_FORM;
  }

  // Check for SitePages before list view pattern (since listView pattern includes SitePages)
  if (url.includes('/SitePages/')) {
    return SharePointPageType.PAGE;
  }

  // Check for document libraries (with both encoded and unencoded spaces)
  if (url.includes('/Shared%20Documents') || url.includes('/Shared Documents')) {
    return SharePointPageType.LIBRARY;
  }

  // Check for list views
  if (SHAREPOINT_URL_PATTERNS.listView.test(url)) {
    return SharePointPageType.LIST_VIEW;
  }

  // Check for other forms (using /Forms/ path)
  if (url.includes('/Forms/')) {
    return SharePointPageType.LIST_FORM;
  }

  // Check for generic pages
  if (url.endsWith('.aspx')) {
    return SharePointPageType.PAGE;
  }

  // Check for other document libraries
  if (url.includes('/Documents')) {
    return SharePointPageType.LIBRARY;
  }

  // Default to site home if at root
  if (urlObj.pathname === '/' || urlObj.pathname.match(/^\/sites\/[^/]+\/?$/)) {
    return SharePointPageType.SITE_HOME;
  }

  return SharePointPageType.UNKNOWN;
}

/**
 * Gets the complete SharePoint context for the current page
 */
export function getSharePointContext(url: string = window.location.href): SharePointContext {
  const isSharePoint = isSharePointPage(url);

  if (!isSharePoint) {
    return {
      isSharePoint: false,
      version: SharePointVersion.UNKNOWN,
      url,
    };
  }

  const version = detectSharePointVersion(url);
  const siteUrl = extractSiteUrl(url);
  const webUrl = extractWebUrl(url);
  const listId = extractListId(url);
  const pageType = detectPageType(url);

  return {
    isSharePoint: true,
    version,
    url,
    siteUrl,
    webUrl,
    listId,
    pageType,
  };
}
