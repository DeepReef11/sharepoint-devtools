/**
 * SharePoint context and detection types
 */

/**
 * SharePoint version type
 */
export enum SharePointVersion {
  CLASSIC = 'classic',
  MODERN = 'modern',
  UNKNOWN = 'unknown',
}

/**
 * SharePoint context information
 */
export interface SharePointContext {
  // Detection
  isSharePoint: boolean;
  version: SharePointVersion;

  // URLs
  url: string;
  siteUrl?: string;
  webUrl?: string;

  // Context IDs
  listId?: string;
  pageId?: string;

  // Page type
  pageType?: SharePointPageType;
}

/**
 * SharePoint page types
 */
export enum SharePointPageType {
  SITE_HOME = 'site_home',
  LIST_VIEW = 'list_view',
  LIST_FORM = 'list_form',
  LIBRARY = 'library',
  PAGE = 'page',
  SITE_SETTINGS = 'site_settings',
  ADMIN = 'admin',
  UNKNOWN = 'unknown',
}

/**
 * SharePoint URL patterns for detection
 */
export const SHAREPOINT_URL_PATTERNS = {
  sharepoint: /\.sharepoint\.com/i,
  sharepointDf: /\.sharepoint-df\.com/i,
  modernPage: /_layouts\/15\/.*\.aspx/i,
  classicPage: /_layouts\/.*\.aspx/i,
  listView: /\/Lists\/|\/SitePages\/|AllItems\.aspx/i,
  siteSettings: /_layouts\/15\/settings\.aspx/i,
};
