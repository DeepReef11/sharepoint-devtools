/**
 * Custom Links Storage Module
 * Manages persistence of user-defined SharePoint links using chrome.storage.sync
 */

import { SharePointLink } from '../links/types';

const STORAGE_KEY = 'customLinks';
const STORAGE_VERSION_KEY = 'customLinksVersion';
const CURRENT_VERSION = '1.0';

/**
 * Custom link storage interface
 */
export interface CustomLinksData {
  version: string;
  links: SharePointLink[];
  lastModified: string;
}

/**
 * Storage result wrapper
 */
export interface StorageResult<T> {
  success: boolean;
  data?: T;
  error?: string;
}

/**
 * Custom Links Storage Manager
 */
export class CustomLinksStorage {
  /**
   * Loads all custom links from storage
   */
  static async loadCustomLinks(): Promise<StorageResult<SharePointLink[]>> {
    try {
      const result = await chrome.storage.sync.get([STORAGE_KEY, STORAGE_VERSION_KEY]);

      if (!result[STORAGE_KEY]) {
        // No custom links yet, return empty array
        return {
          success: true,
          data: [],
        };
      }

      const data = result[STORAGE_KEY] as CustomLinksData;

      // Version check for future migrations
      if (data.version !== CURRENT_VERSION) {
        console.warn(
          `Custom links version mismatch. Expected ${CURRENT_VERSION}, got ${data.version}`
        );
        // Could perform migration here if needed
      }

      return {
        success: true,
        data: data.links || [],
      };
    } catch (error) {
      console.error('Failed to load custom links:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  /**
   * Saves custom links to storage
   */
  static async saveCustomLinks(links: SharePointLink[]): Promise<StorageResult<void>> {
    try {
      const data: CustomLinksData = {
        version: CURRENT_VERSION,
        links,
        lastModified: new Date().toISOString(),
      };

      await chrome.storage.sync.set({
        [STORAGE_KEY]: data,
        [STORAGE_VERSION_KEY]: CURRENT_VERSION,
      });

      return {
        success: true,
      };
    } catch (error) {
      console.error('Failed to save custom links:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  /**
   * Adds a new custom link
   */
  static async addCustomLink(link: SharePointLink): Promise<StorageResult<SharePointLink[]>> {
    const loadResult = await this.loadCustomLinks();
    if (!loadResult.success) {
      return {
        success: false,
        error: loadResult.error,
      };
    }

    const existingLinks = loadResult.data || [];

    // Check for duplicate ID
    if (existingLinks.some((l) => l.id === link.id)) {
      return {
        success: false,
        error: `A link with ID "${link.id}" already exists`,
      };
    }

    const updatedLinks = [...existingLinks, link];
    const saveResult = await this.saveCustomLinks(updatedLinks);

    if (!saveResult.success) {
      return {
        success: false,
        error: saveResult.error,
      };
    }

    return {
      success: true,
      data: updatedLinks,
    };
  }

  /**
   * Updates an existing custom link
   */
  static async updateCustomLink(
    linkId: string,
    updates: Partial<SharePointLink>
  ): Promise<StorageResult<SharePointLink[]>> {
    const loadResult = await this.loadCustomLinks();
    if (!loadResult.success) {
      return {
        success: false,
        error: loadResult.error,
      };
    }

    const existingLinks = loadResult.data || [];
    const linkIndex = existingLinks.findIndex((l) => l.id === linkId);

    if (linkIndex === -1) {
      return {
        success: false,
        error: `Link with ID "${linkId}" not found`,
      };
    }

    const updatedLinks = [...existingLinks];
    updatedLinks[linkIndex] = {
      ...updatedLinks[linkIndex],
      ...updates,
      id: linkId, // Ensure ID cannot be changed
    };

    const saveResult = await this.saveCustomLinks(updatedLinks);

    if (!saveResult.success) {
      return {
        success: false,
        error: saveResult.error,
      };
    }

    return {
      success: true,
      data: updatedLinks,
    };
  }

  /**
   * Deletes a custom link by ID
   */
  static async deleteCustomLink(linkId: string): Promise<StorageResult<SharePointLink[]>> {
    const loadResult = await this.loadCustomLinks();
    if (!loadResult.success) {
      return {
        success: false,
        error: loadResult.error,
      };
    }

    const existingLinks = loadResult.data || [];
    const updatedLinks = existingLinks.filter((l) => l.id !== linkId);

    if (updatedLinks.length === existingLinks.length) {
      return {
        success: false,
        error: `Link with ID "${linkId}" not found`,
      };
    }

    const saveResult = await this.saveCustomLinks(updatedLinks);

    if (!saveResult.success) {
      return {
        success: false,
        error: saveResult.error,
      };
    }

    return {
      success: true,
      data: updatedLinks,
    };
  }

  /**
   * Exports custom links as JSON
   */
  static async exportCustomLinks(): Promise<StorageResult<string>> {
    const loadResult = await this.loadCustomLinks();
    if (!loadResult.success) {
      return {
        success: false,
        error: loadResult.error,
      };
    }

    try {
      const data: CustomLinksData = {
        version: CURRENT_VERSION,
        links: loadResult.data || [],
        lastModified: new Date().toISOString(),
      };

      const json = JSON.stringify(data, null, 2);
      return {
        success: true,
        data: json,
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to serialize links',
      };
    }
  }

  /**
   * Imports custom links from JSON
   */
  static async importCustomLinks(
    jsonData: string,
    mode: 'replace' | 'merge' = 'replace'
  ): Promise<StorageResult<SharePointLink[]>> {
    try {
      const data: CustomLinksData = JSON.parse(jsonData);

      // Validate structure
      if (!data.links || !Array.isArray(data.links)) {
        return {
          success: false,
          error: 'Invalid import format: missing or invalid "links" array',
        };
      }

      // Validate each link
      const validationError = this.validateLinks(data.links);
      if (validationError) {
        return {
          success: false,
          error: validationError,
        };
      }

      let finalLinks = data.links;

      if (mode === 'merge') {
        const loadResult = await this.loadCustomLinks();
        if (!loadResult.success) {
          return {
            success: false,
            error: loadResult.error,
          };
        }

        const existingLinks = loadResult.data || [];
        const existingIds = new Set(existingLinks.map((l) => l.id));

        // Merge: keep existing links and add new ones
        finalLinks = [...existingLinks];

        for (const newLink of data.links) {
          if (!existingIds.has(newLink.id)) {
            finalLinks.push(newLink);
          } else {
            // Update existing link
            const index = finalLinks.findIndex((l) => l.id === newLink.id);
            finalLinks[index] = newLink;
          }
        }
      }

      const saveResult = await this.saveCustomLinks(finalLinks);

      if (!saveResult.success) {
        return {
          success: false,
          error: saveResult.error,
        };
      }

      return {
        success: true,
        data: finalLinks,
      };
    } catch (error) {
      return {
        success: false,
        error:
          error instanceof Error ? `Failed to parse JSON: ${error.message}` : 'Invalid JSON format',
      };
    }
  }

  /**
   * Validates an array of links
   */
  private static validateLinks(links: any[]): string | null {
    for (let i = 0; i < links.length; i++) {
      const link = links[i];

      if (!link.id || typeof link.id !== 'string') {
        return `Link at index ${i} has missing or invalid "id"`;
      }

      if (!link.title || typeof link.title !== 'string') {
        return `Link at index ${i} has missing or invalid "title"`;
      }

      if (!link.urlTemplate || typeof link.urlTemplate !== 'string') {
        return `Link at index ${i} has missing or invalid "urlTemplate"`;
      }

      if (!link.category || typeof link.category !== 'string') {
        return `Link at index ${i} has missing or invalid "category"`;
      }

      if (!link.description || typeof link.description !== 'string') {
        return `Link at index ${i} has missing or invalid "description"`;
      }
    }

    return null;
  }

  /**
   * Clears all custom links
   */
  static async clearCustomLinks(): Promise<StorageResult<void>> {
    try {
      await chrome.storage.sync.remove([STORAGE_KEY, STORAGE_VERSION_KEY]);
      return {
        success: true,
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to clear custom links',
      };
    }
  }

  /**
   * Gets storage usage statistics
   */
  static async getStorageStats(): Promise<StorageResult<{ bytesUsed: number; quota: number }>> {
    try {
      const bytesInUse = await chrome.storage.sync.getBytesInUse(STORAGE_KEY);
      const quota = chrome.storage.sync.QUOTA_BYTES;

      return {
        success: true,
        data: {
          bytesUsed: bytesInUse,
          quota,
        },
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to get storage stats',
      };
    }
  }
}
