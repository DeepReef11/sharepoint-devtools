/**
 * Link Tracker Tests
 */

import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { LinkTracker } from '../src/storage/link-tracker';
import { StorageManager } from '../src/storage/storage-manager';
import { SharePointLink } from '../src/links/types';

// Mock StorageManager
jest.mock('../src/storage/storage-manager');

describe('LinkTracker', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('trackLinkAccess', () => {
    it('should track link access', async () => {
      const link: SharePointLink = {
        id: 'site-settings',
        title: 'Site Settings',
        urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
        category: 'Site Admin' as any,
        description: 'Main site settings page',
      };

      const resolvedUrl = 'https://example.sharepoint.com/_layouts/15/settings.aspx';

      (StorageManager.addRecentLink as any).mockResolvedValue(undefined);

      await LinkTracker.trackLinkAccess(link, resolvedUrl);

      expect(StorageManager.addRecentLink).toHaveBeenCalledWith({
        linkId: 'site-settings',
        title: 'Site Settings',
        url: resolvedUrl,
        category: 'Site Admin',
      });
    });

    it('should not throw on tracking errors', async () => {
      const link: SharePointLink = {
        id: 'site-settings',
        title: 'Site Settings',
        urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
        category: 'Site Admin' as any,
        description: 'Main site settings page',
      };

      (StorageManager.addRecentLink as any).mockRejectedValue(new Error('Storage error'));

      await expect(
        LinkTracker.trackLinkAccess(
          link,
          'https://example.sharepoint.com/_layouts/15/settings.aspx'
        )
      ).resolves.not.toThrow();
    });
  });

  describe('getRecentLinks', () => {
    it('should get recent links', async () => {
      const mockLinks = [
        {
          linkId: 'site-settings',
          title: 'Site Settings',
          url: 'https://example.sharepoint.com/_layouts/15/settings.aspx',
          category: 'Site Admin',
          lastAccessed: Date.now(),
          accessCount: 3,
        },
      ];

      (StorageManager.getRecentLinks as any).mockResolvedValue(mockLinks);

      const links = await LinkTracker.getRecentLinks();
      expect(links).toEqual(mockLinks);
    });

    it('should limit recent links', async () => {
      const mockLinks = Array.from({ length: 10 }, (_, i) => ({
        linkId: `link-${i}`,
        title: `Link ${i}`,
        url: `https://example.com/${i}`,
        category: 'Test',
        lastAccessed: Date.now(),
        accessCount: 1,
      }));

      (StorageManager.getRecentLinks as any).mockResolvedValue(mockLinks);

      const links = await LinkTracker.getRecentLinks(5);
      expect(links).toHaveLength(5);
    });
  });

  describe('getMostAccessedLinks', () => {
    it('should get most accessed links sorted by access count', async () => {
      const mockLinks = [
        {
          linkId: 'link-1',
          title: 'Link 1',
          url: 'https://example.com/1',
          category: 'Test',
          lastAccessed: Date.now(),
          accessCount: 5,
        },
        {
          linkId: 'link-2',
          title: 'Link 2',
          url: 'https://example.com/2',
          category: 'Test',
          lastAccessed: Date.now(),
          accessCount: 10,
        },
        {
          linkId: 'link-3',
          title: 'Link 3',
          url: 'https://example.com/3',
          category: 'Test',
          lastAccessed: Date.now(),
          accessCount: 3,
        },
      ];

      (StorageManager.getRecentLinks as any).mockResolvedValue(mockLinks);

      const links = await LinkTracker.getMostAccessedLinks(2);
      expect(links).toHaveLength(2);
      expect(links[0].linkId).toBe('link-2');
      expect(links[0].accessCount).toBe(10);
      expect(links[1].linkId).toBe('link-1');
      expect(links[1].accessCount).toBe(5);
    });
  });

  describe('clearHistory', () => {
    it('should clear recent links history', async () => {
      (StorageManager.clearRecentLinks as any).mockResolvedValue(undefined);

      await LinkTracker.clearHistory();
      expect(StorageManager.clearRecentLinks).toHaveBeenCalled();
    });
  });

  describe('getLinkStats', () => {
    it('should get stats for a specific link', async () => {
      const mockLinks = [
        {
          linkId: 'site-settings',
          title: 'Site Settings',
          url: 'https://example.sharepoint.com/_layouts/15/settings.aspx',
          category: 'Site Admin',
          lastAccessed: 1234567890,
          accessCount: 5,
        },
      ];

      (StorageManager.getRecentLinks as any).mockResolvedValue(mockLinks);

      const stats = await LinkTracker.getLinkStats('site-settings');
      expect(stats).toEqual({
        accessCount: 5,
        lastAccessed: 1234567890,
      });
    });

    it('should return null for non-existent link', async () => {
      (StorageManager.getRecentLinks as any).mockResolvedValue([]);

      const stats = await LinkTracker.getLinkStats('non-existent');
      expect(stats).toBeNull();
    });
  });
});
