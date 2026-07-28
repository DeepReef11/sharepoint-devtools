/**
 * Multi-tenant support type definitions
 */

/**
 * Represents a SharePoint tenant
 */
export interface Tenant {
  /**
   * The base URL of the tenant (e.g., https://contoso.sharepoint.com)
   */
  tenantUrl: string;

  /**
   * User-friendly display name for the tenant
   */
  displayName?: string;

  /**
   * Last time this tenant was accessed
   */
  lastAccessed: string;

  /**
   * Whether this is the default tenant for the user
   */
  isDefault?: boolean;

  /**
   * Tenant-specific custom links (optional)
   */
  customLinks?: any[];

  /**
   * Additional tenant-specific settings
   */
  settings?: Record<string, any>;
}

/**
 * User preferences for multi-tenant support
 */
export interface TenantPreferences {
  /**
   * Whether to automatically detect tenant from current URL
   * Default: true
   */
  autoDetectTenant: boolean;

  /**
   * Default tenant URL to use when not on a SharePoint page
   */
  defaultTenantUrl?: string;

  /**
   * Maximum number of recent tenants to track
   * Default: 10
   */
  maxRecentTenants?: number;
}

/**
 * Storage structure for tenant data in chrome.storage
 */
export interface TenantStorageData {
  /**
   * List of known tenants
   */
  tenants: Tenant[];

  /**
   * User preferences for multi-tenant behavior
   */
  preferences: TenantPreferences;

  /**
   * Schema version for future migrations
   */
  version: number;
}

/**
 * Cloud environments for SharePoint
 */
export enum SharePointCloud {
  /**
   * *.sharepoint.com — also serves GCC tenants, which are not distinguishable
   * from commercial ones by hostname.
   */
  Commercial = 'commercial',
  /** *.sharepoint.us */
  GCCHigh = 'gcc-high',
  /** *.sharepoint-mil.us */
  DoD = 'dod',
  /** *.sharepoint.cn (operated by 21Vianet) */
  China = 'china',
  /** *.sharepoint-df.com — Microsoft-internal first-release ring */
  Dogfood = 'dogfood',
  Unknown = 'unknown',
}

/**
 * Tenant information with cloud detection
 */
export interface TenantInfo {
  /**
   * The tenant URL
   */
  tenantUrl: string;

  /**
   * The tenant name (extracted from URL)
   * e.g., "contoso" from "contoso.sharepoint.com"
   */
  tenantName: string;

  /**
   * The cloud environment
   */
  cloud: SharePointCloud;

  /**
   * The full hostname
   */
  hostname: string;
}

/**
 * Result of tenant switching operation
 */
export interface TenantSwitchResult {
  success: boolean;
  previousTenantUrl: string | null;
  newTenantUrl: string;
  error?: string;
}
