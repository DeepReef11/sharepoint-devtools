/**
 * Storage Types
 * Defines types for recent and favorite links stored in chrome.storage
 */

/**
 * A link that has been accessed by the user
 */
export interface RecentLink {
  /** Link ID from registry */
  linkId: string;
  /** Display title */
  title: string;
  /** Resolved URL */
  url: string;
  /** Category */
  category: string;
  /** Last accessed timestamp */
  lastAccessed: number;
  /** Number of times accessed */
  accessCount: number;
}

/**
 * A link marked as favorite by the user
 */
export interface FavoriteLink {
  /** Link ID from registry */
  linkId: string;
  /** Display title */
  title: string;
  /** URL template (before resolution) */
  urlTemplate: string;
  /** Category */
  category: string;
  /** When it was favorited */
  favoritedAt: number;
  /** Optional user note */
  note?: string;
}

/**
 * Storage structure for recent links
 */
export interface RecentLinksStorage {
  /** Array of recent links, sorted by lastAccessed (newest first) */
  links: RecentLink[];
  /** Maximum number of recent links to store */
  maxSize: number;
}

/**
 * Storage structure for favorite links
 */
export interface FavoriteLinksStorage {
  /** Map of linkId to favorite link */
  links: { [linkId: string]: FavoriteLink };
}

/**
 * User preferences for recent/favorite links
 */
export interface LinkPreferences {
  /** Maximum recent links to keep */
  maxRecentLinks: number;
  /** Whether to show recent links in quick access */
  showRecent: boolean;
  /** Whether to show favorites in quick access */
  showFavorites: boolean;
  /** Sort order for favorites */
  favoriteSortOrder: 'recent' | 'alphabetical' | 'category';
}

/**
 * Storage keys used in chrome.storage
 */
export enum StorageKey {
  RecentLinks = 'recentLinks',
  FavoriteLinks = 'favoriteLinks',
  Preferences = 'linkPreferences',
}

/**
 * Default preferences
 */
export const DEFAULT_PREFERENCES: LinkPreferences = {
  maxRecentLinks: 10,
  showRecent: true,
  showFavorites: true,
  favoriteSortOrder: 'recent',
};
