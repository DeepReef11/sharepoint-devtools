/**
 * Supported placeholder types for link templates
 */
export enum Placeholder {
  WebUrl = '{webUrl}',
  SiteUrl = '{siteUrl}',
  ListId = '{listId}',
  ListUrl = '{listUrl}',
}

/**
 * Link categories for organizing SharePoint navigation links
 */
export enum LinkCategory {
  SiteAdmin = 'Site Admin',
  Content = 'Content',
  Navigation = 'Navigation',
  Design = 'Design',
  Development = 'Development',
  ModernAdmin = 'Modern Admin',
  ListLibrarySettings = 'List/Library Settings',
  Advanced = 'Advanced',
}

/**
 * Context requirements for a link to be applicable
 */
export interface LinkContext {
  /** Requires user to be in a list/library context */
  requiresList?: boolean;
  /** Requires site collection admin permissions */
  requiresSiteAdmin?: boolean;
  /** Requires tenant admin permissions */
  requiresTenantAdmin?: boolean;
  /** Only applicable in Modern SharePoint */
  modernOnly?: boolean;
  /** Only applicable in Classic SharePoint */
  classicOnly?: boolean;
}

/**
 * Represents a SharePoint navigation link
 */
export interface SharePointLink {
  /** Unique identifier for the link */
  id: string;
  /** Display title for the link */
  title: string;
  /** URL template with placeholders */
  urlTemplate: string;
  /** Category for grouping and filtering */
  category: LinkCategory;
  /** Description of what the link does */
  description: string;
  /** Context requirements for applicability */
  context?: LinkContext;
  /** Keywords for search optimization */
  keywords?: string[];
  /** Display priority (lower number = higher priority) */
  priority?: number;
}

/**
 * Grouped links by category
 */
export interface GroupedLinks {
  [key: string]: SharePointLink[];
}

/**
 * Placeholder values extracted from current context
 */
export interface PlaceholderValues {
  webUrl?: string;
  siteUrl?: string;
  serverUrl?: string;
  tenantAdminUrl?: string;
  listId?: string;
  listUrl?: string;
}
