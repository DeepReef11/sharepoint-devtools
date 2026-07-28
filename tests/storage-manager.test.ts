/**
 * Storage Manager Tests
 * Tests for chrome.storage interactions
 */

import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { StorageManager } from '../src/storage/storage-manager';
import { RecentLink, FavoriteLink } from '../src/storage/types';

// Mock chrome.storage API
const mockStorage: any = {
  local: {
    get: jest.fn(),
    set: jest.fn(),
    getBytesInUse: jest.fn(),
  },
};

(global as any).chrome = {
  storage: mockStorage,
};

describe('StorageManager', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    mockStorage.local.get.mockResolvedValue({});
    mockStorage.local.set.mockResolvedValue(undefined);
    mockStorage.local.getBytesInUse.mockResolvedValue(0);
  });

  describe('Recent Links', () => {
    it('should get recent links from storage', async () => {
      const mockLinks: RecentLink[] = [
        {
          linkId: 'site-settings',
          title: 'Site Settings',
          url: 'https://example.sharepoint.com/_layouts/15/settings.aspx',
          category: 'Site Admin',
          lastAccessed: Date.now(),
          accessCount: 3,
        },
      ];

      mockStorage.local.get.mockResolvedValue({
        recentLinks: { links: mockLinks, maxSize: 10 },
      });

      const links = await StorageManager.getRecentLinks();
      expect(links).toEqual(mockLinks);
    });

    it('should return empty array if no recent links', async () => {
      mockStorage.local.get.mockResolvedValue({});
      const links = await StorageManager.getRecentLinks();
      expect(links).toEqual([]);
    });

    it('should add a new recent link', async () => {
      mockStorage.local.get.mockResolvedValue({
        recentLinks: { links: [], maxSize: 10 },
        linkPreferences: { maxRecentLinks: 10 },
      });

      await StorageManager.addRecentLink({
        linkId: 'site-settings',
        title: 'Site Settings',
        url: 'https://example.sharepoint.com/_layouts/15/settings.aspx',
        category: 'Site Admin',
      });

      expect(mockStorage.local.set).toHaveBeenCalledWith(
        expect.objectContaining({
          recentLinks: expect.objectContaining({
            links: expect.arrayContaining([
              expect.objectContaining({
                linkId: 'site-settings',
                accessCount: 1,
              }),
            ]),
          }),
        })
      );
    });

    it('should update existing recent link and move to front', async () => {
      const existingLink: RecentLink = {
        linkId: 'site-settings',
        title: 'Site Settings',
        url: 'https://example.sharepoint.com/_layouts/15/settings.aspx',
        category: 'Site Admin',
        lastAccessed: Date.now() - 10000,
        accessCount: 2,
      };

      mockStorage.local.get.mockResolvedValue({
        recentLinks: { links: [existingLink], maxSize: 10 },
        linkPreferences: { maxRecentLinks: 10 },
      });

      await StorageManager.addRecentLink({
        linkId: 'site-settings',
        title: 'Site Settings',
        url: 'https://example.sharepoint.com/_layouts/15/settings.aspx',
        category: 'Site Admin',
      });

      expect(mockStorage.local.set).toHaveBeenCalledWith(
        expect.objectContaining({
          recentLinks: expect.objectContaining({
            links: expect.arrayContaining([
              expect.objectContaining({
                linkId: 'site-settings',
                accessCount: 3,
              }),
            ]),
          }),
        })
      );
    });
  });

  describe('Favorite Links', () => {
    it('should get favorite links from storage', async () => {
      const mockFavorites: { [key: string]: FavoriteLink } = {
        'site-settings': {
          linkId: 'site-settings',
          title: 'Site Settings',
          urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
          category: 'Site Admin',
          favoritedAt: Date.now(),
        },
      };

      mockStorage.local.get.mockResolvedValue({
        favoriteLinks: { links: mockFavorites },
      });

      const favorites = await StorageManager.getFavoriteLinks();
      expect(favorites).toHaveLength(1);
      expect(favorites[0].linkId).toBe('site-settings');
    });

    it('should check if a link is favorited', async () => {
      const mockFavorites: { [key: string]: FavoriteLink } = {
        'site-settings': {
          linkId: 'site-settings',
          title: 'Site Settings',
          urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
          category: 'Site Admin',
          favoritedAt: Date.now(),
        },
      };

      mockStorage.local.get.mockResolvedValue({
        favoriteLinks: { links: mockFavorites },
      });

      const isFav = await StorageManager.isFavorite('site-settings');
      expect(isFav).toBe(true);

      const isNotFav = await StorageManager.isFavorite('other-link');
      expect(isNotFav).toBe(false);
    });

    it('should add a favorite', async () => {
      mockStorage.local.get.mockResolvedValue({
        favoriteLinks: { links: {} },
      });

      await StorageManager.addFavorite({
        linkId: 'site-settings',
        title: 'Site Settings',
        urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
        category: 'Site Admin',
      });

      expect(mockStorage.local.set).toHaveBeenCalledWith(
        expect.objectContaining({
          favoriteLinks: expect.objectContaining({
            links: expect.objectContaining({
              'site-settings': expect.objectContaining({
                linkId: 'site-settings',
              }),
            }),
          }),
        })
      );
    });

    it('should remove a favorite', async () => {
      const mockFavorites: { [key: string]: FavoriteLink } = {
        'site-settings': {
          linkId: 'site-settings',
          title: 'Site Settings',
          urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
          category: 'Site Admin',
          favoritedAt: Date.now(),
        },
      };

      mockStorage.local.get.mockResolvedValue({
        favoriteLinks: { links: mockFavorites },
      });

      await StorageManager.removeFavorite('site-settings');

      expect(mockStorage.local.set).toHaveBeenCalledWith(
        expect.objectContaining({
          favoriteLinks: expect.objectContaining({
            links: {},
          }),
        })
      );
    });

    it('should toggle favorite status', async () => {
      // First call - not favorited
      mockStorage.local.get.mockResolvedValueOnce({
        favoriteLinks: { links: {} },
      });

      const added = await StorageManager.toggleFavorite({
        linkId: 'site-settings',
        title: 'Site Settings',
        urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
        category: 'Site Admin',
      });

      expect(added).toBe(true);

      // Second call - already favorited
      const mockFavorites: { [key: string]: FavoriteLink } = {
        'site-settings': {
          linkId: 'site-settings',
          title: 'Site Settings',
          urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
          category: 'Site Admin',
          favoritedAt: Date.now(),
        },
      };

      mockStorage.local.get.mockResolvedValueOnce({
        favoriteLinks: { links: mockFavorites },
      });

      const removed = await StorageManager.toggleFavorite({
        linkId: 'site-settings',
        title: 'Site Settings',
        urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
        category: 'Site Admin',
      });

      expect(removed).toBe(false);
    });
  });

  describe('Preferences', () => {
    it('should get preferences with defaults', async () => {
      mockStorage.local.get.mockResolvedValue({});
      const prefs = await StorageManager.getPreferences();
      expect(prefs.maxRecentLinks).toBe(10);
      expect(prefs.showRecent).toBe(true);
      expect(prefs.showFavorites).toBe(true);
    });

    it('should save preferences', async () => {
      mockStorage.local.get.mockResolvedValue({
        linkPreferences: {
          maxRecentLinks: 10,
          showRecent: true,
          showFavorites: true,
          favoriteSortOrder: 'recent',
        },
      });

      await StorageManager.savePreferences({ maxRecentLinks: 20 });

      expect(mockStorage.local.set).toHaveBeenCalledWith(
        expect.objectContaining({
          linkPreferences: expect.objectContaining({
            maxRecentLinks: 20,
          }),
        })
      );
    });
  });

  describe('Storage Stats', () => {
    it('should get storage statistics', async () => {
      mockStorage.local.get.mockImplementation((key) => {
        if (key === 'recentLinks') {
          return Promise.resolve({
            recentLinks: { links: [{ linkId: '1' }, { linkId: '2' }] },
          });
        }
        if (key === 'favoriteLinks') {
          return Promise.resolve({
            favoriteLinks: { links: { '1': {}, '2': {}, '3': {} } },
          });
        }
        return Promise.resolve({});
      });

      mockStorage.local.getBytesInUse.mockResolvedValue(1024);

      const stats = await StorageManager.getStorageStats();
      expect(stats.recentCount).toBe(2);
      expect(stats.favoriteCount).toBe(3);
      expect(stats.bytesUsed).toBe(1024);
    });
  });
});
