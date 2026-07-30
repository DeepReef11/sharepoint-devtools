/**
 * Fuzzy Search Tests
 * Validates the fuzzy search functionality using Fuse.js
 */

import { describe, it, expect, beforeEach } from '@jest/globals';
import { createFuzzySearch, FuzzySearchService } from '../src/links/fuzzy-search';
import { SHAREPOINT_LINKS } from '../src/links/registry';
import { SharePointLink, LinkCategory } from '../src/links/types';

describe('FuzzySearchService', () => {
  let testLinks: SharePointLink[];
  let searchService: FuzzySearchService;

  beforeEach(() => {
    testLinks = [
      {
        id: 'site-settings',
        title: 'Site Settings',
        urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
        category: LinkCategory.SiteAdmin,
        description: 'Main site settings page',
        keywords: ['settings', 'configuration', 'admin'],
        priority: 1,
      },
      {
        id: 'site-permissions',
        title: 'Site Permissions',
        urlTemplate: '{webUrl}/_layouts/15/user.aspx',
        category: LinkCategory.SiteAdmin,
        description: 'Manage site permissions',
        keywords: ['permissions', 'security', 'access'],
        priority: 2,
      },
      {
        id: 'recycle-bin',
        title: 'Recycle Bin',
        urlTemplate: '{webUrl}/_layouts/15/RecycleBin.aspx',
        category: LinkCategory.Content,
        description: 'Restore deleted items',
        keywords: ['recycle', 'deleted', 'restore'],
        priority: 3,
      },
      {
        id: 'list-settings',
        title: 'List Settings',
        urlTemplate: '{listUrl}/_layouts/15/listedit.aspx?List={listId}',
        category: LinkCategory.ListLibrarySettings,
        description: 'Configure list or library',
        keywords: ['list', 'library', 'settings'],
      },
      {
        id: 'content-types',
        title: 'Site Content Types',
        urlTemplate: '{webUrl}/_layouts/15/mngctype.aspx',
        category: LinkCategory.Content,
        description: 'Manage content types',
        keywords: ['content', 'types', 'schema'],
      },
    ];

    searchService = createFuzzySearch(testLinks);
  });

  describe('Basic Search', () => {
    it('should find results for exact matches', () => {
      const results = searchService.search('settings');
      expect(results.length).toBeGreaterThan(0);
      expect(results.some((r) => r.item.id === 'site-settings')).toBe(true);
    });

    it('should return empty array for empty query', () => {
      const results = searchService.search('');
      expect(results).toEqual([]);
    });

    it('should return empty array for whitespace-only query', () => {
      const results = searchService.search('   ');
      expect(results).toEqual([]);
    });

    it('should return no results for non-existent terms', () => {
      const results = searchService.search('xyz123nonexistent');
      expect(results).toEqual([]);
    });

    // "xyz123nonexistent" above is long and shares no run of letters with
    // anything, so it clears the bar however loose the matcher is. The case
    // that actually breaks is an ordinary short word that happens to overlap
    // one: "zebra" contains "bra", which sits inside "library".
    it('rejects an unrelated word that overlaps a keyword', () => {
      const results = searchService.search('zebra');
      expect(results).toEqual([]);
    });
  });

  describe('Match precision against the shipped registry', () => {
    // The registry is the data users search, and it is what made this visible:
    // "zebra" returned 15 of its links, the best scoring 0.03 — better than
    // most real matches.
    const registry = () => createFuzzySearch(SHAREPOINT_LINKS);

    it.each(['zebra', 'pizza', 'banana', 'qwerty', 'wombat'])(
      'returns nothing for %p',
      (query) => {
        expect(registry().search(query, { limit: 20 })).toEqual([]);
      }
    );

    // Tightening the threshold must not cost real matching. Each of these is
    // something a user types: an exact word, two typos, and two prefixes.
    it.each([
      ['recycle', 'Recycle Bin'],
      ['recyle', 'Recycle Bin'],
      ['permisions', 'Site Permissions'],
      ['colum', 'List Columns'],
      ['versioning', 'Versioning Settings'],
      ['term store', 'Term Store Management'],
    ])('still finds %p', (query, expected) => {
      const titles = registry()
        .search(query, { limit: 20 })
        .map((r) => r.item.title);
      expect(titles).toContain(expected);
    });
  });

  describe('Fuzzy Matching', () => {
    it('should handle typos in search queries', () => {
      const results = searchService.search('permisions'); // Typo: "permisions" instead of "permissions"
      expect(results.some((r) => r.item.id === 'site-permissions')).toBe(true);
    });

    it('should handle partial matches', () => {
      const results = searchService.search('recy'); // Partial of "recycle"
      expect(results.some((r) => r.item.id === 'recycle-bin')).toBe(true);
    });

    it('should match with case insensitivity', () => {
      const results = searchService.search('SETTINGS');
      expect(results.length).toBeGreaterThan(0);
      expect(results.some((r) => r.item.id === 'site-settings')).toBe(true);
    });
  });

  describe('Multi-field Search', () => {
    it('should search in title field', () => {
      const results = searchService.search('Recycle');
      expect(results.some((r) => r.item.id === 'recycle-bin')).toBe(true);
    });

    it('should search in description field', () => {
      const results = searchService.search('restore'); // In description
      expect(results.some((r) => r.item.id === 'recycle-bin')).toBe(true);
    });

    it('should search in keywords field', () => {
      const results = searchService.search('security'); // In keywords
      expect(results.some((r) => r.item.id === 'site-permissions')).toBe(true);
    });

    it('should search in category field', () => {
      const results = searchService.search('admin');
      expect(results.length).toBeGreaterThan(0);
    });
  });

  describe('Search Options', () => {
    it('should respect limit option', () => {
      const results = searchService.search('settings', { limit: 2 });
      expect(results.length).toBeLessThanOrEqual(2);
    });

    it('should include scores when requested', () => {
      const results = searchService.search('settings', { includeScore: true });
      results.forEach((result) => {
        expect(result.score).toBeDefined();
        expect(typeof result.score).toBe('number');
      });
    });

    it('should include matches when requested', () => {
      const results = searchService.search('settings', { includeMatches: true });
      results.forEach((result) => {
        expect(result.matches).toBeDefined();
      });
    });
  });

  describe('Search Ranking', () => {
    it('should rank results by relevance (lower score is better)', () => {
      const results = searchService.search('site');
      for (let i = 0; i < results.length - 1; i++) {
        expect(results[i].score!).toBeLessThanOrEqual(results[i + 1].score!);
      }
    });

    it('should prioritize title matches over other fields', () => {
      const results = searchService.search('settings');
      // The first result should be one with "settings" in the title
      expect(results[0].item.title.toLowerCase()).toContain('settings');
    });
  });

  describe('Index Management', () => {
    it('should update search index with setLinks', () => {
      const newLinks: SharePointLink[] = [
        {
          id: 'new-link',
          title: 'New Link',
          urlTemplate: '{webUrl}/new',
          category: LinkCategory.Content,
          description: 'A new link',
        },
      ];

      searchService.setLinks(newLinks);
      const results = searchService.search('new');
      expect(results.some((r) => r.item.id === 'new-link')).toBe(true);

      // Old links should not be found
      const oldResults = searchService.search('settings');
      expect(oldResults.length).toBe(0);
    });

    it('should add a link to the index', () => {
      const newLink: SharePointLink = {
        id: 'added-link',
        title: 'Added Link',
        urlTemplate: '{webUrl}/added',
        category: LinkCategory.Content,
        description: 'An added link',
      };

      searchService.addLink(newLink);
      const results = searchService.search('added');
      expect(results.some((r) => r.item.id === 'added-link')).toBe(true);
    });

    it('should remove a link from the index', () => {
      searchService.removeLink('site-settings');
      const results = searchService.search('site settings');
      expect(results.some((r) => r.item.id === 'site-settings')).toBe(false);
    });
  });

  describe('Options Management', () => {
    it('should return current options', () => {
      const options = searchService.getOptions();
      expect(options).toBeDefined();
      expect(options.threshold).toBeDefined();
    });

    it('should update options', () => {
      searchService.updateOptions({ threshold: 0.2 });
      const options = searchService.getOptions();
      expect(options.threshold).toBe(0.2);
    });
  });

  describe('Factory Function', () => {
    it('should create a new FuzzySearchService instance', () => {
      const service = createFuzzySearch(testLinks);
      expect(service).toBeInstanceOf(FuzzySearchService);
    });

    it('should accept custom options', () => {
      const service = createFuzzySearch(testLinks, { threshold: 0.1 });
      const options = service.getOptions();
      expect(options.threshold).toBe(0.1);
    });
  });
});
