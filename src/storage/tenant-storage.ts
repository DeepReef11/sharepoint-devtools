/**
 * Tenant storage manager for chrome.storage
 */

import { Tenant, TenantPreferences, TenantStorageData } from '../types/tenant';

/**
 * Storage key for tenant data in chrome.storage
 */
const STORAGE_KEY = 'sharepoint_quicknav_tenants';

/**
 * Current storage schema version
 */
const STORAGE_VERSION = 1;

/**
 * Default preferences for multi-tenant support
 */
const DEFAULT_PREFERENCES: TenantPreferences = {
  autoDetectTenant: true,
  maxRecentTenants: 10,
};

/**
 * Manages tenant data persistence using chrome.storage.sync
 */
export class TenantStorage {
  /**
   * Loads all tenant data from storage
   */
  async loadData(): Promise<TenantStorageData> {
    try {
      const result = await chrome.storage.sync.get(STORAGE_KEY);

      if (result[STORAGE_KEY]) {
        const data = result[STORAGE_KEY] as TenantStorageData;

        // Handle version migrations if needed
        if (data.version !== STORAGE_VERSION) {
          return this.migrateData(data);
        }

        return data;
      }

      // Return default data if nothing stored
      return this.getDefaultData();
    } catch (error) {
      console.error('Failed to load tenant data from storage:', error);
      return this.getDefaultData();
    }
  }

  /**
   * Saves tenant data to storage
   */
  async saveData(data: TenantStorageData): Promise<void> {
    try {
      await chrome.storage.sync.set({ [STORAGE_KEY]: data });
    } catch (error) {
      console.error('Failed to save tenant data to storage:', error);
      throw error;
    }
  }

  /**
   * Gets a specific tenant by URL
   */
  async getTenant(tenantUrl: string): Promise<Tenant | null> {
    const data = await this.loadData();
    return data.tenants.find((t) => t.tenantUrl === tenantUrl) || null;
  }

  /**
   * Adds or updates a tenant
   */
  async saveTenant(tenant: Tenant): Promise<void> {
    const data = await this.loadData();

    const existingIndex = data.tenants.findIndex((t) => t.tenantUrl === tenant.tenantUrl);

    if (existingIndex >= 0) {
      // Update existing tenant
      data.tenants[existingIndex] = {
        ...data.tenants[existingIndex],
        ...tenant,
        lastAccessed: new Date().toISOString(),
      };
    } else {
      // Add new tenant
      data.tenants.push({
        ...tenant,
        lastAccessed: new Date().toISOString(),
      });

      // Limit number of recent tenants
      const maxTenants = data.preferences.maxRecentTenants || 10;
      if (data.tenants.length > maxTenants) {
        // Sort by lastAccessed and keep only the most recent
        data.tenants.sort(
          (a, b) => new Date(b.lastAccessed).getTime() - new Date(a.lastAccessed).getTime()
        );
        data.tenants = data.tenants.slice(0, maxTenants);
      }
    }

    await this.saveData(data);
  }

  /**
   * Removes a tenant from storage
   */
  async removeTenant(tenantUrl: string): Promise<void> {
    const data = await this.loadData();
    data.tenants = data.tenants.filter((t) => t.tenantUrl !== tenantUrl);

    // Clear default if removing the default tenant
    if (data.preferences.defaultTenantUrl === tenantUrl) {
      data.preferences.defaultTenantUrl = undefined;
    }

    await this.saveData(data);
  }

  /**
   * Gets all tenants sorted by last accessed
   */
  async getAllTenants(): Promise<Tenant[]> {
    const data = await this.loadData();
    return [...data.tenants].sort(
      (a, b) => new Date(b.lastAccessed).getTime() - new Date(a.lastAccessed).getTime()
    );
  }

  /**
   * Gets the default tenant
   */
  async getDefaultTenant(): Promise<Tenant | null> {
    const data = await this.loadData();

    if (data.preferences.defaultTenantUrl) {
      const tenant = data.tenants.find((t) => t.tenantUrl === data.preferences.defaultTenantUrl);
      if (tenant) return tenant;
    }

    // Find tenant marked as default
    const defaultTenant = data.tenants.find((t) => t.isDefault);
    if (defaultTenant) return defaultTenant;

    // Return most recently accessed tenant
    if (data.tenants.length > 0) {
      const sorted = [...data.tenants].sort(
        (a, b) => new Date(b.lastAccessed).getTime() - new Date(a.lastAccessed).getTime()
      );
      return sorted[0];
    }

    return null;
  }

  /**
   * Sets the default tenant
   */
  async setDefaultTenant(tenantUrl: string): Promise<void> {
    const data = await this.loadData();

    // Clear all default flags
    data.tenants.forEach((t) => {
      t.isDefault = false;
    });

    // Set new default
    const tenant = data.tenants.find((t) => t.tenantUrl === tenantUrl);
    if (tenant) {
      tenant.isDefault = true;
    }

    data.preferences.defaultTenantUrl = tenantUrl;
    await this.saveData(data);
  }

  /**
   * Gets user preferences
   */
  async getPreferences(): Promise<TenantPreferences> {
    const data = await this.loadData();
    return data.preferences;
  }

  /**
   * Updates user preferences
   */
  async updatePreferences(preferences: Partial<TenantPreferences>): Promise<void> {
    const data = await this.loadData();
    data.preferences = {
      ...data.preferences,
      ...preferences,
    };
    await this.saveData(data);
  }

  /**
   * Updates last accessed time for a tenant
   */
  async touchTenant(tenantUrl: string): Promise<void> {
    const data = await this.loadData();
    const tenant = data.tenants.find((t) => t.tenantUrl === tenantUrl);

    if (tenant) {
      tenant.lastAccessed = new Date().toISOString();
      await this.saveData(data);
    }
  }

  /**
   * Clears all tenant data (for testing or reset)
   */
  async clearAll(): Promise<void> {
    try {
      await chrome.storage.sync.remove(STORAGE_KEY);
    } catch (error) {
      console.error('Failed to clear tenant data:', error);
      throw error;
    }
  }

  /**
   * Returns default storage data
   */
  private getDefaultData(): TenantStorageData {
    return {
      tenants: [],
      preferences: { ...DEFAULT_PREFERENCES },
      version: STORAGE_VERSION,
    };
  }

  /**
   * Migrates data from old schema versions to current
   */
  private migrateData(data: TenantStorageData): TenantStorageData {
    // Future: Add migration logic here when schema changes
    console.warn('Data migration needed from version', data.version, 'to', STORAGE_VERSION);

    // For now, just update version
    return {
      ...data,
      version: STORAGE_VERSION,
    };
  }
}

/**
 * Singleton instance of TenantStorage
 */
export const tenantStorage = new TenantStorage();
