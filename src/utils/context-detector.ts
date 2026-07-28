/**
 * SharePoint Context Detector
 * Parses URLs and page state to extract SharePoint context information
 * SharePoint context detection with multi-tenant support.
 */

import {
  SharePointContext,
  SharePointPageType,
  SharePointVersion,
  PageContext,
} from '../types/sharepoint-context';
import { tenantManager } from '../managers/tenant-manager';
import { Tenant } from '../types/tenant';
import { isSharePointHostname } from './sharepoint-clouds';

/**
 * Detects SharePoint context from a given URL
 * @param url - The URL to parse (defaults to current window.location.href)
 * @returns SharePoint context information
 */
export function detectSharePointContext(url?: string): SharePointContext {
  const currentUrl = url || (typeof window !== 'undefined' ? window.location.href : '');

  // Check if this is a SharePoint URL
  const isSharePointPage = isSharePointUrl(currentUrl);

  if (!isSharePointPage) {
    return createEmptyContext(currentUrl);
  }

  // Parse URL components
  const urlObj = new URL(currentUrl);
  const tenantUrl = extractTenantUrl(urlObj);
  const siteUrl = extractSiteUrl(urlObj);
  const webUrl = extractWebUrl(urlObj);
  const listId = extractListId(urlObj);
  const listTitle = extractListTitle(urlObj);
  const pageType = detectPageType(urlObj, listId);
  const version = detectSharePointVersion(urlObj);

  // Extract tenant ID from domain
  const tenantId = urlObj.hostname.split('.')[0];

  return {
    isSharePoint: true,
    version,
    pageContext: PageContext.Unknown, // Simplified detector - use context module for detailed detection
    pageType,
    tenantUrl,
    tenantId,
    siteUrl,
    webUrl,
    listId,
    listTitle,
    urlComponents: null,
    listContext: null,
    isSharePointPage,
    currentUrl,
    detectedAt: new Date(),
  };
}

/**
 * Checks if a URL is a SharePoint URL
 */
export function isSharePointUrl(url: string): boolean {
  if (!url) return false;

  try {
    const urlObj = new URL(url);
    const hostname = urlObj.hostname.toLowerCase();

    // Check for SharePoint Online patterns
    return isSharePointHostname(hostname);
  } catch {
    return false;
  }
}

/**
 * Extracts tenant URL (root of SharePoint tenant)
 */
function extractTenantUrl(urlObj: URL): string | null {
  const protocol = urlObj.protocol;
  const hostname = urlObj.hostname;

  // For SharePoint Online, tenant URL is just protocol + hostname
  if (hostname.includes('.sharepoint.')) {
    return `${protocol}//${hostname}`;
  }

  return null;
}

/**
 * Extracts site collection URL
 * Examples:
 * - https://contoso.sharepoint.com/sites/mysite
 * - https://contoso.sharepoint.com (root site)
 */
function extractSiteUrl(urlObj: URL): string | null {
  const tenantUrl = extractTenantUrl(urlObj);
  if (!tenantUrl) return null;

  const pathname = urlObj.pathname;

  // Check for /sites/ or /teams/ pattern
  const sitesMatch = pathname.match(/^\/(sites|teams)\/([^/]+)/);
  if (sitesMatch) {
    return `${tenantUrl}${sitesMatch[0]}`;
  }

  // Root site collection
  if (pathname === '/' || pathname.startsWith('/_')) {
    return tenantUrl;
  }

  // Check for personal OneDrive site
  if (pathname.startsWith('/personal/')) {
    const personalMatch = pathname.match(/^\/personal\/[^/]+/);
    if (personalMatch) {
      return `${tenantUrl}${personalMatch[0]}`;
    }
  }

  // Default to tenant URL for root site
  return tenantUrl;
}

/**
 * Extracts web/subsite URL
 * For now, same as site URL (subsites detection requires additional logic)
 */
function extractWebUrl(urlObj: URL): string | null {
  // TODO: Implement subsite detection in future iteration
  // For now, web URL is the same as site URL
  return extractSiteUrl(urlObj);
}

/**
 * Extracts list/library GUID from URL
 * Can be in query string (?List={guid}) or path (/Lists/ListName)
 */
function extractListId(urlObj: URL): string | null {
  // Check query parameters for List parameter
  const listParam = urlObj.searchParams.get('List');
  if (listParam) {
    // Remove curly braces if present
    return listParam.replace(/[{}]/g, '');
  }

  // Check for RootFolder parameter which might contain list ID
  const rootFolder = urlObj.searchParams.get('RootFolder');
  if (rootFolder) {
    const guidMatch = rootFolder.match(
      /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i
    );
    if (guidMatch) {
      return guidMatch[0];
    }
  }

  // Check URL path for GUID patterns
  const guidMatch = urlObj.pathname.match(
    /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i
  );
  if (guidMatch) {
    return guidMatch[0];
  }

  return null;
}

/**
 * Extracts list/library title from URL path
 */
function extractListTitle(urlObj: URL): string | null {
  const pathname = urlObj.pathname;

  // Check for /Lists/ListName or /LibraryName pattern
  const listsMatch = pathname.match(/\/Lists\/([^/]+)/);
  if (listsMatch) {
    return decodeURIComponent(listsMatch[1]);
  }

  // Check for library patterns (common library names)
  const libraryMatch = pathname.match(/\/(Documents|SiteAssets|SitePages|Style Library)(?:\/|$)/);
  if (libraryMatch) {
    return libraryMatch[1];
  }

  return null;
}

/**
 * Detects the page type based on URL patterns
 */
function detectPageType(urlObj: URL, listId: string | null): SharePointPageType {
  const pathname = urlObj.pathname.toLowerCase();

  // List or library pages
  if (listId || pathname.includes('/lists/') || pathname.includes('/forms/')) {
    // Check if it's a library (document library patterns)
    if (
      pathname.includes('/documents/') ||
      pathname.includes('/siteassets/') ||
      pathname.includes('/sitepages/') ||
      pathname.includes('/style library/')
    ) {
      return SharePointPageType.Library;
    }
    return SharePointPageType.List;
  }

  // Site pages
  if (pathname.includes('/sitepages/') || pathname.endsWith('.aspx')) {
    return SharePointPageType.Page;
  }

  // Site home or other site pages
  if (
    pathname.includes('/_layouts/') ||
    pathname === '/' ||
    pathname.includes('/sites/') ||
    pathname.includes('/teams/')
  ) {
    return SharePointPageType.Site;
  }

  return SharePointPageType.Unknown;
}

/**
 * Detects SharePoint version (Modern vs Classic)
 */
function detectSharePointVersion(urlObj: URL): SharePointVersion {
  const pathname = urlObj.pathname.toLowerCase();

  // Classic UI indicators
  if (pathname.includes('/_layouts/15/') || pathname.includes('/_layouts/')) {
    return SharePointVersion.Classic;
  }

  // Modern UI indicators
  if (
    pathname.includes('/sitepages/') ||
    urlObj.searchParams.has('view') ||
    urlObj.searchParams.has('viewid')
  ) {
    // Additional check: Classic list forms
    if (pathname.includes('/forms/') && pathname.endsWith('.aspx')) {
      return SharePointVersion.Classic;
    }
    return SharePointVersion.Modern;
  }

  // Try to detect from DOM if we're in a browser context
  if (typeof document !== 'undefined') {
    // Modern UI has specific meta tags or CSS classes
    const isModern =
      document.querySelector('div[data-automationid="modernPageContainer"]') !== null ||
      document.querySelector('.od-SuiteNav') !== null;

    if (isModern) {
      return SharePointVersion.Modern;
    }

    // Classic UI indicators
    const isClassic =
      document.getElementById('s4-workspace') !== null ||
      document.querySelector('.ms-core-pageTitle') !== null;

    if (isClassic) {
      return SharePointVersion.Classic;
    }
  }

  return SharePointVersion.Unknown;
}

/**
 * Creates an empty context for non-SharePoint URLs
 */
function createEmptyContext(url: string): SharePointContext {
  return {
    isSharePoint: false,
    version: SharePointVersion.Unknown,
    pageContext: PageContext.Unknown,
    pageType: SharePointPageType.Unknown,
    tenantUrl: null,
    tenantId: null,
    siteUrl: null,
    webUrl: null,
    listId: null,
    listTitle: null,
    urlComponents: null,
    listContext: null,
    isSharePointPage: false,
    currentUrl: url,
    detectedAt: new Date(),
  };
}

/**
 * Gets the current SharePoint context
 * Convenience function for browser context
 */
export function getCurrentContext(): SharePointContext {
  if (typeof window === 'undefined') {
    throw new Error('getCurrentContext can only be called in a browser context');
  }

  return detectSharePointContext(window.location.href);
}

/**
 * Detects SharePoint context with tenant management integration
 *
 * This function combines context detection with automatic tenant registration
 * when auto-detection is enabled.
 *
 * @param url - The URL to parse (defaults to current window.location.href)
 * @returns SharePoint context information with tenant data
 */
export async function detectContextWithTenant(
  url?: string
): Promise<{ context: SharePointContext; tenant: Tenant | null }> {
  const context = detectSharePointContext(url);

  if (!context.isSharePointPage || !context.tenantUrl) {
    return { context, tenant: null };
  }

  // Check if auto-detection is enabled
  const autoDetect = await tenantManager.isAutoDetectEnabled();

  // Get or create tenant (auto-register if enabled)
  const tenant = await tenantManager.getOrCreateTenant(context.tenantUrl, autoDetect);

  return { context, tenant };
}

/**
 * Gets the current context with tenant information
 * Convenience function for browser context with tenant management
 */
export async function getCurrentContextWithTenant(): Promise<{
  context: SharePointContext;
  tenant: Tenant | null;
}> {
  if (typeof window === 'undefined') {
    throw new Error('getCurrentContextWithTenant can only be called in a browser context');
  }

  return detectContextWithTenant(window.location.href);
}

/**
 * Validates if the current context is on a known tenant
 *
 * @param context - SharePoint context to validate
 * @returns true if the tenant is registered, false otherwise
 */
export async function isKnownTenant(context: SharePointContext): Promise<boolean> {
  if (!context.tenantUrl) {
    return false;
  }

  const tenant = await tenantManager.getCurrentTenant();
  return tenant !== null && tenant.tenantUrl === context.tenantUrl;
}
