/**
 * Caching layer for SharePoint REST API responses
 * Reduces network requests and improves performance
 */

/* global sessionStorage */

export interface CacheEntry<T> {
  data: T;
  timestamp: number;
  ttl: number;
}

export interface CacheOptions {
  /**
   * Time to live in milliseconds
   * Default: 5 minutes (300000ms)
   */
  ttl?: number;

  /**
   * Use session storage instead of memory cache
   * Survives page refreshes within the same session
   * Default: false
   */
  useSessionStorage?: boolean;
}

/**
 * In-memory cache for API responses
 */
class ApiCache {
  private memoryCache = new Map<string, CacheEntry<any>>();
  private readonly DEFAULT_TTL = 5 * 60 * 1000; // 5 minutes
  private readonly SESSION_STORAGE_PREFIX = 'spqn_cache_';

  /**
   * Get cached data
   */
  get<T>(key: string, useSessionStorage = false): T | null {
    // Try session storage first if requested
    if (useSessionStorage) {
      const sessionData = this.getFromSessionStorage<T>(key);
      if (sessionData !== null) {
        return sessionData;
      }
    }

    // Try memory cache
    const entry = this.memoryCache.get(key);
    if (!entry) {
      return null;
    }

    // Check if entry has expired
    const now = Date.now();
    if (now - entry.timestamp > entry.ttl) {
      this.memoryCache.delete(key);
      return null;
    }

    return entry.data;
  }

  /**
   * Set cached data
   */
  set<T>(key: string, data: T, options: CacheOptions = {}): void {
    const ttl = options.ttl ?? this.DEFAULT_TTL;
    const entry: CacheEntry<T> = {
      data,
      timestamp: Date.now(),
      ttl,
    };

    // Store in memory cache
    this.memoryCache.set(key, entry);

    // Store in session storage if requested
    if (options.useSessionStorage) {
      this.setInSessionStorage(key, entry);
    }
  }

  /**
   * Clear specific cache entry
   */
  delete(key: string, useSessionStorage = false): void {
    this.memoryCache.delete(key);

    if (useSessionStorage) {
      try {
        sessionStorage.removeItem(this.SESSION_STORAGE_PREFIX + key);
      } catch (error) {
        console.warn('Failed to delete from session storage:', error);
      }
    }
  }

  /**
   * Clear all cache entries
   */
  clear(clearSessionStorage = false): void {
    this.memoryCache.clear();

    if (clearSessionStorage) {
      try {
        const keysToRemove: string[] = [];
        for (let i = 0; i < sessionStorage.length; i++) {
          const key = sessionStorage.key(i);
          if (key?.startsWith(this.SESSION_STORAGE_PREFIX)) {
            keysToRemove.push(key);
          }
        }
        keysToRemove.forEach((key) => sessionStorage.removeItem(key));
      } catch (error) {
        console.warn('Failed to clear session storage:', error);
      }
    }
  }

  /**
   * Get cache statistics
   */
  getStats(): { size: number; keys: string[] } {
    return {
      size: this.memoryCache.size,
      keys: Array.from(this.memoryCache.keys()),
    };
  }

  /**
   * Get data from session storage
   */
  private getFromSessionStorage<T>(key: string): T | null {
    try {
      const item = sessionStorage.getItem(this.SESSION_STORAGE_PREFIX + key);
      if (!item) {
        return null;
      }

      const entry: CacheEntry<T> = JSON.parse(item);
      const now = Date.now();

      // Check if expired
      if (now - entry.timestamp > entry.ttl) {
        sessionStorage.removeItem(this.SESSION_STORAGE_PREFIX + key);
        return null;
      }

      return entry.data;
    } catch (error) {
      console.warn('Failed to read from session storage:', error);
      return null;
    }
  }

  /**
   * Store data in session storage
   */
  private setInSessionStorage<T>(key: string, entry: CacheEntry<T>): void {
    try {
      sessionStorage.setItem(this.SESSION_STORAGE_PREFIX + key, JSON.stringify(entry));
    } catch (error) {
      // Session storage might be full or disabled
      console.warn('Failed to write to session storage:', error);
    }
  }
}

// Export singleton instance
export const apiCache = new ApiCache();

/**
 * Generate cache key from URL and parameters
 */
export function generateCacheKey(url: string, params?: Record<string, any>): string {
  const baseKey = url.replace(/[^a-zA-Z0-9]/g, '_');
  if (!params || Object.keys(params).length === 0) {
    return baseKey;
  }

  const paramStr = Object.entries(params)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([k, v]) => `${k}=${v}`)
    .join('_');

  return `${baseKey}_${paramStr}`;
}
