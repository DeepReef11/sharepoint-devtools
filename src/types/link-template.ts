/**
 * Link Template Types
 * Defines the structure for SharePoint link templates with placeholders
 * and applicability rules for the Link Resolution Engine
 */

import { SharePointPageType, SharePointVersion } from './sharepoint-context';

/**
 * Link categories for organization
 */
export enum LinkCategory {
  SiteAdmin = 'Site Admin',
  Content = 'Content',
  Navigation = 'Navigation',
  Design = 'Design',
  Development = 'Development',
  ModernAdmin = 'Modern Admin',
  ListSettings = 'List/Library Settings',
  Advanced = 'Advanced',
}

/**
 * Placeholder types supported in link templates
 */
export enum PlaceholderType {
  TenantUrl = 'tenantUrl',
  SiteUrl = 'siteUrl',
  WebUrl = 'webUrl',
  ListId = 'listId',
  ListTitle = 'listTitle',
}

/**
 * Applicability rules for a link template
 * Determines when a link should be shown based on current context
 */
export interface LinkApplicability {
  /** Required page types (empty = all) */
  pageTypes?: SharePointPageType[];

  /** Required SharePoint versions (empty = all) */
  versions?: SharePointVersion[];

  /** Required context fields that must be present */
  requiredContext?: PlaceholderType[];

  /** Minimum permission level needed (for future implementation) */
  minimumPermission?: 'view' | 'edit' | 'admin';
}

/**
 * Link template with placeholders and applicability rules
 */
export interface LinkTemplate {
  /** Unique identifier */
  id: string;

  /** Display title */
  title: string;

  /** URL template with placeholders (e.g., {webUrl}/_layouts/15/settings.aspx) */
  urlTemplate: string;

  /** Category for grouping */
  category: LinkCategory;

  /** Description for display/search */
  description?: string;

  /** Keywords for fuzzy search */
  keywords?: string[];

  /** Applicability rules */
  applicability: LinkApplicability;

  /** Whether this link opens in a new tab */
  openInNewTab?: boolean;
}

/**
 * Resolved link ready for display/navigation
 */
export interface ResolvedLink {
  /** Reference to original template */
  templateId: string;

  /** Display title */
  title: string;

  /** Fully resolved URL */
  url: string;

  /** Category */
  category: LinkCategory;

  /** Description */
  description?: string;

  /** Keywords */
  keywords?: string[];

  /** Whether to open in new tab */
  openInNewTab: boolean;

  /** Whether the URL passed validation */
  isValid: boolean;

  /** Validation errors if any */
  validationErrors?: string[];
}

/**
 * Link registry containing all link templates
 */
export interface LinkRegistry {
  /** All link templates */
  templates: LinkTemplate[];

  /** Version of the registry */
  version: string;
}
