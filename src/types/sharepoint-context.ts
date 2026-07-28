/**
 * SharePoint Context Detection Types
 * Defines data structures for SharePoint site context information
 */

/**
 * SharePoint version/type enumeration
 */
export enum SharePointVersion {
  /** Modern SharePoint Online experience */
  Modern = 'Modern',
  /** Classic SharePoint experience */
  Classic = 'Classic',
  /** Unable to determine version */
  Unknown = 'Unknown',
}

/**
 * SharePoint page types that can be detected
 * Used by the Link Resolution Engine to determine applicable links
 */
export enum SharePointPageType {
  Site = 'site',
  List = 'list',
  Library = 'library',
  Page = 'page',
  Unknown = 'unknown',
}

/**
 * SharePoint page context types (detailed)
 */
export enum PageContext {
  /** Site home page */
  SiteHome = 'SiteHome',
  /** Document library */
  DocumentLibrary = 'DocumentLibrary',
  /** List */
  List = 'List',
  /** List item view/edit */
  ListItem = 'ListItem',
  /** Site settings page */
  SiteSettings = 'SiteSettings',
  /** Site contents page */
  SiteContents = 'SiteContents',
  /** Modern page */
  ModernPage = 'ModernPage',
  /** Wiki page */
  WikiPage = 'WikiPage',
  /** Web part page */
  WebPartPage = 'WebPartPage',
  /** Unknown page type */
  Unknown = 'Unknown',
}

/**
 * Parsed SharePoint URL components
 */
export interface SharePointUrlComponents {
  /** Full URL that was parsed */
  fullUrl: string;
  /** Protocol (http/https) */
  protocol: string;
  /** SharePoint tenant domain (e.g., contoso.sharepoint.com) */
  domain: string;
  /** Root site collection URL */
  siteCollectionUrl: string;
  /** Web (site) URL - may be subsite */
  webUrl: string;
  /** Relative path within the web */
  relativePath: string;
  /** Query string parameters */
  queryParams: Map<string, string>;
}

/**
 * List/Library context information
 */
export interface ListContext {
  /** List/Library GUID */
  listId: string | null;
  /** List/Library internal name */
  listName: string | null;
  /** List/Library title (display name) */
  listTitle: string | null;
  /** Item ID if viewing/editing a specific item */
  itemId: string | null;
  /** View ID if a specific view is active */
  viewId: string | null;
}

/**
 * Complete SharePoint context information
 * Extracted from the current URL and page state
 */
export interface SharePointContext {
  /** Whether the current page is a SharePoint site */
  isSharePoint: boolean;

  /** SharePoint version/type */
  version: SharePointVersion;

  /** Current page context type (detailed) */
  pageContext: PageContext;

  /** Current page type (simplified for Link Resolution Engine) */
  pageType: SharePointPageType;

  /** Full tenant URL (e.g., https://contoso.sharepoint.com) */
  tenantUrl: string | null;

  /** Tenant identifier (extracted from domain) */
  tenantId: string | null;

  /** Site collection URL (e.g., https://contoso.sharepoint.com/sites/mysite) */
  siteUrl: string | null;

  /** Web/subsite URL (e.g., https://contoso.sharepoint.com/sites/mysite/subweb) */
  webUrl: string | null;

  /** List or library GUID (if on a list/library page) */
  listId: string | null;

  /** List or library title (if available) */
  listTitle: string | null;

  /** Parsed URL components */
  urlComponents: SharePointUrlComponents | null;

  /** List/Library context (if applicable) */
  listContext: ListContext | null;

  /** Whether the current page is a SharePoint page */
  isSharePointPage: boolean;

  /** Original URL being parsed */
  currentUrl: string;

  /** Timestamp when context was detected */
  detectedAt: Date;
}

/**
 * Result of context detection with potential errors
 */
export interface ContextDetectionResult {
  /** Successfully detected context */
  success: boolean;
  /** The detected context (if successful) */
  context: SharePointContext | null;
  /** Error message (if unsuccessful) */
  error: string | null;
}

/**
 * Partial context for testing or incomplete detection
 */
export type PartialSharePointContext = Partial<SharePointContext>;
