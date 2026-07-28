/**
 * Storage Manager
 * Handles all interactions with chrome.storage for recent and favorite links
 */

import {
  RecentLink,
  FavoriteLink,
  RecentLinksStorage,
  FavoriteLinksStorage,
  LinkPreferences,
  StorageKey,
  DEFAULT_PREFERENCES,
} from './types';

/**
 * Manages storage operations for recent and favorite links
 */
export class StorageManager {
  /**
   * Get recent links from storage
   */
  static async getRecentLinks(): Promise<RecentLink[]> {
    try {
      const result = await chrome.storage.local.get(StorageKey.RecentLinks);
      const storage = result[StorageKey.RecentLinks] as RecentLinksStorage | undefined;
      return storage?.links || [];
    } catch (error) {
      console.error('Error getting recent links:', error);
      return [];
    }
  }

  /**
   * Save recent links to storage
   */
  static async saveRecentLinks(links: RecentLink[]): Promise<void> {
    try {
      const preferences = await this.getPreferences();
      const storage: RecentLinksStorage = {
        links: links.slice(0, preferences.maxRecentLinks),
        maxSize: preferences.maxRecentLinks,
      };
      await chrome.storage.local.set({ [StorageKey.RecentLinks]: storage });
    } catch (error) {
      console.error('Error saving recent links:', error);
      throw error;
    }
  }

  /**
   * Add or update a recent link
   */
  static async addRecentLink(
    link: Omit<RecentLink, 'lastAccessed' | 'accessCount'>
  ): Promise<void> {
    try {
      const recentLinks = await this.getRecentLinks();
      const existingIndex = recentLinks.findIndex((l) => l.linkId === link.linkId);

      if (existingIndex >= 0) {
        // Update existing link
        const existing = recentLinks[existingIndex];
        recentLinks[existingIndex] = {
          ...link,
          lastAccessed: Date.now(),
          accessCount: existing.accessCount + 1,
        };
        // Move to front
        const updated = recentLinks.splice(existingIndex, 1)[0];
        recentLinks.unshift(updated);
      } else {
        // Add new link at the beginning
        recentLinks.unshift({
          ...link,
          lastAccessed: Date.now(),
          accessCount: 1,
        });
      }

      await this.saveRecentLinks(recentLinks);
    } catch (error) {
      console.error('Error adding recent link:', error);
      throw error;
    }
  }

  /**
   * Clear all recent links
   */
  static async clearRecentLinks(): Promise<void> {
    try {
      await chrome.storage.local.set({ [StorageKey.RecentLinks]: { links: [], maxSize: 10 } });
    } catch (error) {
      console.error('Error clearing recent links:', error);
      throw error;
    }
  }

  /**
   * Get favorite links from storage
   */
  static async getFavoriteLinks(): Promise<FavoriteLink[]> {
    try {
      const result = await chrome.storage.local.get(StorageKey.FavoriteLinks);
      const storage = result[StorageKey.FavoriteLinks] as FavoriteLinksStorage | undefined;
      return storage?.links ? Object.values(storage.links) : [];
    } catch (error) {
      console.error('Error getting favorite links:', error);
      return [];
    }
  }

  /**
   * Check if a link is favorited
   */
  static async isFavorite(linkId: string): Promise<boolean> {
    try {
      const result = await chrome.storage.local.get(StorageKey.FavoriteLinks);
      const storage = result[StorageKey.FavoriteLinks] as FavoriteLinksStorage | undefined;
      return storage?.links?.[linkId] !== undefined;
    } catch (error) {
      console.error('Error checking favorite:', error);
      return false;
    }
  }

  /**
   * Add a link to favorites
   */
  static async addFavorite(link: Omit<FavoriteLink, 'favoritedAt'>): Promise<void> {
    try {
      const result = await chrome.storage.local.get(StorageKey.FavoriteLinks);
      const storage: FavoriteLinksStorage = (result[
        StorageKey.FavoriteLinks
      ] as FavoriteLinksStorage) || { links: {} };

      storage.links[link.linkId] = {
        ...link,
        favoritedAt: Date.now(),
      };

      await chrome.storage.local.set({ [StorageKey.FavoriteLinks]: storage });
    } catch (error) {
      console.error('Error adding favorite:', error);
      throw error;
    }
  }

  /**
   * Remove a link from favorites
   */
  static async removeFavorite(linkId: string): Promise<void> {
    try {
      const result = await chrome.storage.local.get(StorageKey.FavoriteLinks);
      const storage = result[StorageKey.FavoriteLinks] as FavoriteLinksStorage | undefined;

      if (storage?.links?.[linkId]) {
        delete storage.links[linkId];
        await chrome.storage.local.set({ [StorageKey.FavoriteLinks]: storage });
      }
    } catch (error) {
      console.error('Error removing favorite:', error);
      throw error;
    }
  }

  /**
   * Toggle favorite status of a link
   */
  static async toggleFavorite(link: Omit<FavoriteLink, 'favoritedAt'>): Promise<boolean> {
    try {
      const isFav = await this.isFavorite(link.linkId);
      if (isFav) {
        await this.removeFavorite(link.linkId);
        return false;
      } else {
        await this.addFavorite(link);
        return true;
      }
    } catch (error) {
      console.error('Error toggling favorite:', error);
      throw error;
    }
  }

  /**
   * Clear all favorites
   */
  static async clearFavorites(): Promise<void> {
    try {
      await chrome.storage.local.set({ [StorageKey.FavoriteLinks]: { links: {} } });
    } catch (error) {
      console.error('Error clearing favorites:', error);
      throw error;
    }
  }

  /**
   * Get user preferences
   */
  static async getPreferences(): Promise<LinkPreferences> {
    try {
      const result = await chrome.storage.local.get(StorageKey.Preferences);
      return (result[StorageKey.Preferences] as LinkPreferences) || DEFAULT_PREFERENCES;
    } catch (error) {
      console.error('Error getting preferences:', error);
      return DEFAULT_PREFERENCES;
    }
  }

  /**
   * Save user preferences
   */
  static async savePreferences(preferences: Partial<LinkPreferences>): Promise<void> {
    try {
      const current = await this.getPreferences();
      const updated = { ...current, ...preferences };
      await chrome.storage.local.set({ [StorageKey.Preferences]: updated });
    } catch (error) {
      console.error('Error saving preferences:', error);
      throw error;
    }
  }

  /**
   * Get storage usage statistics
   */
  static async getStorageStats(): Promise<{
    recentCount: number;
    favoriteCount: number;
    bytesUsed: number;
  }> {
    try {
      const [recent, favorites, bytesInUse] = await Promise.all([
        this.getRecentLinks(),
        this.getFavoriteLinks(),
        chrome.storage.local.getBytesInUse(),
      ]);

      return {
        recentCount: recent.length,
        favoriteCount: favorites.length,
        bytesUsed: bytesInUse || 0,
      };
    } catch (error) {
      console.error('Error getting storage stats:', error);
      return { recentCount: 0, favoriteCount: 0, bytesUsed: 0 };
    }
  }

  /**
   * Export all data (for backup/debugging)
   */
  static async exportData(): Promise<{
    recent: RecentLink[];
    favorites: FavoriteLink[];
    preferences: LinkPreferences;
  }> {
    const [recent, favorites, preferences] = await Promise.all([
      this.getRecentLinks(),
      this.getFavoriteLinks(),
      this.getPreferences(),
    ]);

    return { recent, favorites, preferences };
  }

  /**
   * Import data (for restore)
   */
  static async importData(data: {
    recent?: RecentLink[];
    favorites?: FavoriteLink[];
    preferences?: LinkPreferences;
  }): Promise<void> {
    try {
      if (data.recent) {
        await this.saveRecentLinks(data.recent);
      }
      if (data.favorites) {
        const storage: FavoriteLinksStorage = {
          links: data.favorites.reduce(
            (acc, fav) => {
              acc[fav.linkId] = fav;
              return acc;
            },
            {} as { [key: string]: FavoriteLink }
          ),
        };
        await chrome.storage.local.set({ [StorageKey.FavoriteLinks]: storage });
      }
      if (data.preferences) {
        await this.savePreferences(data.preferences);
      }
    } catch (error) {
      console.error('Error importing data:', error);
      throw error;
    }
  }
}
