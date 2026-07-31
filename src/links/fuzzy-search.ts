import Fuse, { IFuseOptions, FuseResultMatch } from 'fuse.js';
import { SharePointLink } from './types';

/**
 * Configuration for fuzzy search
 */
export interface FuzzySearchOptions {
  /** Threshold for matching (0.0 = perfect match, 1.0 = match anything) */
  threshold?: number;
  /** Maximum number of results to return */
  limit?: number;
  /** Whether to include score in results */
  includeScore?: boolean;
  /** Whether to include matches metadata */
  includeMatches?: boolean;
}

/**
 * Search result with score and match information
 */
export interface SearchResult {
  /** The matched link */
  item: SharePointLink;
  /** Match score (lower is better) */
  score?: number;
  /** Matched field indices */
  matches?: readonly FuseResultMatch[];
}

/**
 * Default Fuse.js options optimized for SharePoint link search
 */
const DEFAULT_FUSE_OPTIONS: IFuseOptions<SharePointLink> = {
  // Matching threshold (0.0 = perfect, 1.0 = match anything).
  //
  // 0.4 was too loose to combine with ignoreLocation below. Dropping the
  // distance penalty lets Fuse take its best matching window anywhere in a
  // field, so a query needed only ~60% of its characters to line up somewhere
  // to clear 0.4: "zebra" matched "Li(brar)y" and came back scoring 0.03,
  // better than most genuine matches, filling the palette with 15 unrelated
  // results. Measured across the shipped registry, every value from 0.35 down
  // to 0.20 rejects all of zebra/pizza/banana/qwerty while leaving typo and
  // prefix matching untouched; 0.3 sits in the middle of that band.
  threshold: 0.3,

  // Location of pattern in string (0 = start)
  location: 0,

  // How much distance affects score
  distance: 100,

  // Minimum characters that must be matched
  minMatchCharLength: 2,

  // Search multiple fields with different weights
  keys: [
    {
      name: 'title',
      weight: 2.0, // Title is most important
    },
    {
      name: 'keywords',
      weight: 1.5, // Keywords are also very relevant
    },
    {
      name: 'description',
      weight: 1.0, // Description has normal weight
    },
    {
      name: 'category',
      weight: 0.8, // Category is less important
    },
  ],

  // Return scores and matches for ranking
  includeScore: true,
  includeMatches: true,

  // Use extended search for better results
  useExtendedSearch: false,

  // Ignore location when matching
  ignoreLocation: true,

  // Ignore field-length norm when calculating score
  ignoreFieldNorm: false,

  // Field length norm weight
  fieldNormWeight: 1,
};

/**
 * FuzzySearchService provides fuzzy search functionality for SharePoint links
 * using the Fuse.js library
 */
export class FuzzySearchService {
  private fuse: Fuse<SharePointLink>;
  private options: IFuseOptions<SharePointLink>;
  private links: SharePointLink[];

  /**
   * Creates a new fuzzy search service
   * @param links - Links to search through
   * @param customOptions - Custom Fuse.js options (merged with defaults)
   */
  constructor(links: SharePointLink[], customOptions?: Partial<IFuseOptions<SharePointLink>>) {
    this.options = {
      ...DEFAULT_FUSE_OPTIONS,
      ...customOptions,
    };
    this.links = [...links];
    this.fuse = new Fuse(this.links, this.options);
  }

  /**
   * Performs fuzzy search on the links
   * @param query - Search query string
   * @param options - Search options
   * @returns Array of search results with scores and matches
   */
  search(query: string, options?: FuzzySearchOptions): SearchResult[] {
    if (!query || query.trim().length === 0) {
      return [];
    }

    const limit = options?.limit;
    const fuseResults = this.fuse.search(query, limit ? { limit } : undefined);

    return fuseResults.map((result) => ({
      item: result.item,
      score: options?.includeScore !== false ? result.score : undefined,
      matches: options?.includeMatches !== false ? result.matches : undefined,
    }));
  }

  /**
   * Updates the search index with new links
   * @param links - New links to index
   */
  setLinks(links: SharePointLink[]): void {
    this.links = [...links];
    this.fuse.setCollection(this.links);
  }

  /**
   * Adds a link to the search index
   * @param link - Link to add
   */
  addLink(link: SharePointLink): void {
    this.links.push(link);
    this.fuse.setCollection(this.links);
  }

  /**
   * Removes a link from the search index
   * @param linkId - ID of the link to remove
   */
  removeLink(linkId: string): void {
    this.links = this.links.filter((link) => link.id !== linkId);
    this.fuse.setCollection(this.links);
  }

  /**
   * Gets the current Fuse.js options
   * @returns Current search options
   */
  getOptions(): IFuseOptions<SharePointLink> {
    return { ...this.options };
  }

  /**
   * Updates the search options
   * @param newOptions - New options to merge with current options
   */
  updateOptions(newOptions: Partial<IFuseOptions<SharePointLink>>): void {
    this.options = {
      ...this.options,
      ...newOptions,
    };
    // Recreate the Fuse instance with new options
    this.fuse = new Fuse(this.links, this.options);
  }
}

/**
 * Creates a fuzzy search service instance
 * @param links - Links to search
 * @param options - Custom Fuse.js options
 * @returns Configured fuzzy search service
 */
export function createFuzzySearch(
  links: SharePointLink[],
  options?: Partial<IFuseOptions<SharePointLink>>
): FuzzySearchService {
  return new FuzzySearchService(links, options);
}
