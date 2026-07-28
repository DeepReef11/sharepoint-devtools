/**
 * Link Tracker
 * Service for tracking link access and managing recent links
 */

import { StorageManager } from './storage-manager';
import { SharePointLink } from '../links/types';

/**
 * Tracks link access and manages recent links history
 */
export class LinkTracker {
  /**
   * Track that a link was accessed
   * This should be called whenever a user clicks/opens a link
   */
  static async trackLinkAccess(link: SharePointLink, resolvedUrl: string): Promise<void> {
    try {
      await StorageManager.addRecentLink({
        linkId: link.id,
        title: link.title,
        url: resolvedUrl,
        category: link.category,
      });
    } catch (error) {
      console.error('Error tracking link access:', error);
      // Don't throw - tracking failures shouldn't break navigation
    }
  }

  /**
   * Get recently accessed links
   * @param limit Maximum number of links to return
   */
  static async getRecentLinks(limit?: number): Promise<
    Array<{
      linkId: string;
      title: string;
      url: string;
      category: string;
      lastAccessed: number;
      accessCount: number;
    }>
  > {
    try {
      const links = await StorageManager.getRecentLinks();
      return limit ? links.slice(0, limit) : links;
    } catch (error) {
      console.error('Error getting recent links:', error);
      return [];
    }
  }

  /**
   * Get most frequently accessed links
   * @param limit Maximum number of links to return
   */
  static async getMostAccessedLinks(limit: number = 5): Promise<
    Array<{
      linkId: string;
      title: string;
      url: string;
      category: string;
      accessCount: number;
    }>
  > {
    try {
      const links = await StorageManager.getRecentLinks();
      return links
        .sort((a, b) => b.accessCount - a.accessCount)
        .slice(0, limit)
        .map((link) => ({
          linkId: link.linkId,
          title: link.title,
          url: link.url,
          category: link.category,
          accessCount: link.accessCount,
        }));
    } catch (error) {
      console.error('Error getting most accessed links:', error);
      return [];
    }
  }

  /**
   * Clear recent links history
   */
  static async clearHistory(): Promise<void> {
    try {
      await StorageManager.clearRecentLinks();
    } catch (error) {
      console.error('Error clearing history:', error);
      throw error;
    }
  }

  /**
   * Get access statistics for a specific link
   */
  static async getLinkStats(linkId: string): Promise<{
    accessCount: number;
    lastAccessed: number | null;
  } | null> {
    try {
      const links = await StorageManager.getRecentLinks();
      const link = links.find((l) => l.linkId === linkId);
      if (link) {
        return {
          accessCount: link.accessCount,
          lastAccessed: link.lastAccessed,
        };
      }
      return null;
    } catch (error) {
      console.error('Error getting link stats:', error);
      return null;
    }
  }
}
