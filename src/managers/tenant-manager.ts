/**
 * Tenant manager for multi-tenant support
 * Issue #13: Multi-Tenant Support
 */

import { Tenant, TenantInfo, SharePointCloud, TenantSwitchResult } from '../types/tenant';
import { tenantStorage } from '../storage/tenant-storage';

/**
 * Manages tenant detection, switching, and storage
 */
export class TenantManager {
  private currentTenantUrl: string | null = null;

  /**
   * Detects tenant information from a URL
   */
  detectTenantFromUrl(url: string): TenantInfo | null {
    try {
      const urlObj = new URL(url);
      const hostname = urlObj.hostname.toLowerCase();

      // Check if it's a SharePoint URL
      if (!this.isSharePointUrl(hostname)) {
        return null;
      }

      const tenantUrl = `${urlObj.protocol}//${hostname}`;
      const tenantName = this.extractTenantName(hostname);
      const cloud = this.detectCloud(hostname);

      return {
        tenantUrl,
        tenantName,
        cloud,
        hostname,
      };
    } catch (error) {
      console.error('Failed to detect tenant from URL:', error);
      return null;
    }
  }

  /**
   * Checks if a hostname is a SharePoint URL
   */
  private isSharePointUrl(hostname: string): boolean {
    return (
      hostname.includes('.sharepoint.com') ||
      hostname.includes('.sharepoint-df.com') ||
      hostname.includes('.sharepoint.us') ||
      hostname.includes('.sharepoint.cn')
    );
  }

  /**
   * Extracts tenant name from hostname
   * e.g., "contoso" from "contoso.sharepoint.com"
   */
  private extractTenantName(hostname: string): string {
    const parts = hostname.split('.');

    if (parts.length >= 2) {
      return parts[0];
    }

    return hostname;
  }

  /**
   * Detects the SharePoint cloud environment
   */
  private detectCloud(hostname: string): SharePointCloud {
    if (hostname.includes('.sharepoint.com')) {
      return SharePointCloud.Commercial;
    }
    if (hostname.includes('.sharepoint-df.com')) {
      return SharePointCloud.GCCHigh;
    }
    if (hostname.includes('.sharepoint.us')) {
      return SharePointCloud.GCC;
    }
    if (hostname.includes('.sharepoint.cn')) {
      return SharePointCloud.China;
    }

    return SharePointCloud.Unknown;
  }

  /**
   * Gets or creates a tenant from a URL
   * If auto-detect is enabled, automatically registers new tenants
   */
  async getOrCreateTenant(url: string, autoRegister: boolean = true): Promise<Tenant | null> {
    const tenantInfo = this.detectTenantFromUrl(url);

    if (!tenantInfo) {
      return null;
    }

    // Check if tenant already exists
    let tenant = await tenantStorage.getTenant(tenantInfo.tenantUrl);

    if (!tenant && autoRegister) {
      // Create new tenant entry
      tenant = {
        tenantUrl: tenantInfo.tenantUrl,
        displayName: this.generateDisplayName(tenantInfo),
        lastAccessed: new Date().toISOString(),
      };

      await tenantStorage.saveTenant(tenant);
    }

    // Update current tenant
    if (tenant) {
      this.currentTenantUrl = tenant.tenantUrl;
      await tenantStorage.touchTenant(tenant.tenantUrl);
    }

    return tenant;
  }

  /**
   * Generates a display name for a tenant
   */
  private generateDisplayName(tenantInfo: TenantInfo): string {
    const cloudSuffix =
      tenantInfo.cloud !== SharePointCloud.Commercial ? ` (${tenantInfo.cloud})` : '';

    return `${tenantInfo.tenantName}${cloudSuffix}`;
  }

  /**
   * Gets the current active tenant
   */
  async getCurrentTenant(): Promise<Tenant | null> {
    if (this.currentTenantUrl) {
      return await tenantStorage.getTenant(this.currentTenantUrl);
    }

    // Fall back to default tenant
    return await tenantStorage.getDefaultTenant();
  }

  /**
   * Switches to a different tenant
   */
  async switchTenant(tenantUrl: string): Promise<TenantSwitchResult> {
    const previousTenantUrl = this.currentTenantUrl;

    try {
      const tenant = await tenantStorage.getTenant(tenantUrl);

      if (!tenant) {
        return {
          success: false,
          previousTenantUrl,
          newTenantUrl: tenantUrl,
          error: 'Tenant not found',
        };
      }

      this.currentTenantUrl = tenantUrl;
      await tenantStorage.touchTenant(tenantUrl);

      return {
        success: true,
        previousTenantUrl,
        newTenantUrl: tenantUrl,
      };
    } catch (error) {
      return {
        success: false,
        previousTenantUrl,
        newTenantUrl: tenantUrl,
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  /**
   * Gets all known tenants
   */
  async getAllTenants(): Promise<Tenant[]> {
    return await tenantStorage.getAllTenants();
  }

  /**
   * Updates tenant information
   */
  async updateTenant(tenantUrl: string, updates: Partial<Tenant>): Promise<void> {
    const tenant = await tenantStorage.getTenant(tenantUrl);

    if (!tenant) {
      throw new Error(`Tenant not found: ${tenantUrl}`);
    }

    const updatedTenant: Tenant = {
      ...tenant,
      ...updates,
      // Don't allow changing the tenantUrl
      tenantUrl: tenant.tenantUrl,
    };

    await tenantStorage.saveTenant(updatedTenant);
  }

  /**
   * Removes a tenant
   */
  async removeTenant(tenantUrl: string): Promise<void> {
    await tenantStorage.removeTenant(tenantUrl);

    // Clear current tenant if it was removed
    if (this.currentTenantUrl === tenantUrl) {
      this.currentTenantUrl = null;
    }
  }

  /**
   * Sets a tenant as default
   */
  async setDefaultTenant(tenantUrl: string): Promise<void> {
    await tenantStorage.setDefaultTenant(tenantUrl);
  }

  /**
   * Gets the default tenant
   */
  async getDefaultTenant(): Promise<Tenant | null> {
    return await tenantStorage.getDefaultTenant();
  }

  /**
   * Checks if auto-detection is enabled
   */
  async isAutoDetectEnabled(): Promise<boolean> {
    const preferences = await tenantStorage.getPreferences();
    return preferences.autoDetectTenant;
  }

  /**
   * Sets auto-detection preference
   */
  async setAutoDetect(enabled: boolean): Promise<void> {
    await tenantStorage.updatePreferences({ autoDetectTenant: enabled });
  }

  /**
   * Validates if a URL belongs to a known tenant
   */
  async validateTenantUrl(url: string): Promise<boolean> {
    const tenantInfo = this.detectTenantFromUrl(url);

    if (!tenantInfo) {
      return false;
    }

    const tenant = await tenantStorage.getTenant(tenantInfo.tenantUrl);
    return tenant !== null;
  }

  /**
   * Gets tenant-specific settings
   */
  async getTenantSettings(tenantUrl: string): Promise<Record<string, any>> {
    const tenant = await tenantStorage.getTenant(tenantUrl);
    return tenant?.settings || {};
  }

  /**
   * Updates tenant-specific settings
   */
  async updateTenantSettings(tenantUrl: string, settings: Record<string, any>): Promise<void> {
    const tenant = await tenantStorage.getTenant(tenantUrl);

    if (!tenant) {
      throw new Error(`Tenant not found: ${tenantUrl}`);
    }

    tenant.settings = {
      ...tenant.settings,
      ...settings,
    };

    await tenantStorage.saveTenant(tenant);
  }
}

/**
 * Singleton instance of TenantManager
 */
export const tenantManager = new TenantManager();
