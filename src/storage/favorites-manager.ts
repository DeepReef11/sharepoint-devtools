/**
 * Favorites Manager
 * Service for managing favorite links
 */

import { StorageManager } from './storage-manager';
import { FavoriteLink } from './types';
import { SharePointLink } from '../links/types';

/**
 * Manages favorite links functionality
 */
export class FavoritesManager {
  /**
   * Get all favorite links
   * @param sortOrder Optional sort order (overrides user preference)
   */
  static async getFavorites(
    sortOrder?: 'recent' | 'alphabetical' | 'category'
  ): Promise<FavoriteLink[]> {
    try {
      const favorites = await StorageManager.getFavoriteLinks();
      const preferences = await StorageManager.getPreferences();
      const order = sortOrder || preferences.favoriteSortOrder;

      return this.sortFavorites(favorites, order);
    } catch (error) {
      console.error('Error getting favorites:', error);
      return [];
    }
  }

  /**
   * Sort favorites based on the specified order
   */
  private static sortFavorites(
    favorites: FavoriteLink[],
    order: 'recent' | 'alphabetical' | 'category'
  ): FavoriteLink[] {
    const sorted = [...favorites];

    switch (order) {
      case 'recent':
        sorted.sort((a, b) => b.favoritedAt - a.favoritedAt);
        break;
      case 'alphabetical':
        sorted.sort((a, b) => a.title.localeCompare(b.title));
        break;
      case 'category':
        sorted.sort((a, b) => {
          const catCompare = a.category.localeCompare(b.category);
          if (catCompare !== 0) return catCompare;
          return a.title.localeCompare(b.title);
        });
        break;
    }

    return sorted;
  }

  /**
   * Check if a link is favorited
   */
  static async isFavorite(linkId: string): Promise<boolean> {
    try {
      return await StorageManager.isFavorite(linkId);
    } catch (error) {
      console.error('Error checking favorite:', error);
      return false;
    }
  }

  /**
   * Add a link to favorites
   */
  static async addFavorite(link: SharePointLink, note?: string): Promise<void> {
    try {
      await StorageManager.addFavorite({
        linkId: link.id,
        title: link.title,
        urlTemplate: link.urlTemplate,
        category: link.category,
        note,
      });
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
      await StorageManager.removeFavorite(linkId);
    } catch (error) {
      console.error('Error removing favorite:', error);
      throw error;
    }
  }

  /**
   * Toggle favorite status and return new state
   * @returns true if now favorited, false if unfavorited
   */
  static async toggleFavorite(link: SharePointLink): Promise<boolean> {
    try {
      return await StorageManager.toggleFavorite({
        linkId: link.id,
        title: link.title,
        urlTemplate: link.urlTemplate,
        category: link.category,
      });
    } catch (error) {
      console.error('Error toggling favorite:', error);
      throw error;
    }
  }

  /**
   * Update note for a favorite
   */
  static async updateNote(linkId: string, note: string): Promise<void> {
    try {
      const favorites = await StorageManager.getFavoriteLinks();
      const favorite = favorites.find((f) => f.linkId === linkId);

      if (favorite) {
        await StorageManager.addFavorite({
          ...favorite,
          note,
        });
      }
    } catch (error) {
      console.error('Error updating note:', error);
      throw error;
    }
  }

  /**
   * Get favorites grouped by category
   */
  static async getFavoritesByCategory(): Promise<{ [category: string]: FavoriteLink[] }> {
    try {
      const favorites = await this.getFavorites();
      const grouped: { [category: string]: FavoriteLink[] } = {};

      for (const favorite of favorites) {
        if (!grouped[favorite.category]) {
          grouped[favorite.category] = [];
        }
        grouped[favorite.category].push(favorite);
      }

      return grouped;
    } catch (error) {
      console.error('Error getting favorites by category:', error);
      return {};
    }
  }

  /**
   * Clear all favorites
   */
  static async clearAll(): Promise<void> {
    try {
      await StorageManager.clearFavorites();
    } catch (error) {
      console.error('Error clearing favorites:', error);
      throw error;
    }
  }

  /**
   * Export favorites as JSON
   */
  static async exportFavorites(): Promise<string> {
    try {
      const favorites = await this.getFavorites();
      return JSON.stringify(favorites, null, 2);
    } catch (error) {
      console.error('Error exporting favorites:', error);
      throw error;
    }
  }

  /**
   * Import favorites from JSON
   */
  static async importFavorites(jsonString: string): Promise<number> {
    try {
      const imported = JSON.parse(jsonString) as FavoriteLink[];
      if (!Array.isArray(imported)) {
        throw new Error('Invalid format: expected array of favorites');
      }

      let count = 0;
      for (const fav of imported) {
        if (fav.linkId && fav.title && fav.urlTemplate && fav.category) {
          await StorageManager.addFavorite(fav);
          count++;
        }
      }

      return count;
    } catch (error) {
      console.error('Error importing favorites:', error);
      throw error;
    }
  }

  /**
   * Get favorite count
   */
  static async getCount(): Promise<number> {
    try {
      const favorites = await StorageManager.getFavoriteLinks();
      return favorites.length;
    } catch (error) {
      console.error('Error getting favorite count:', error);
      return 0;
    }
  }
}
