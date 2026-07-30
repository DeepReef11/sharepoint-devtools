import { SharePointLink, LinkCategory, PlaceholderValues, GroupedLinks } from './types';
import { SHAREPOINT_LINKS } from './registry';
import { LinkTemplateEngine } from './template-engine';
import { LinkTracker, FavoritesManager } from '../storage';
import { logError, safely } from '../utils/error-handling';

/**
 * Manages SharePoint links: filtering, resolving, and grouping
 */
export class LinkManager {
  private links: SharePointLink[];
  private baseLinks: SharePointLink[];
  private customLinkIds = new Set<string>();

  constructor(links: SharePointLink[] = SHAREPOINT_LINKS) {
    this.links = links;
    this.baseLinks = links;
  }

  /**
   * Merges the user's own links in alongside the built-in registry.
   *
   * Applied on top of the base registry rather than the current set, so calling
   * it again after the user edits or deletes a link replaces the previous set
   * instead of accumulating. A custom link sharing an id with a built-in wins —
   * overriding a built-in destination is a reasonable thing to want, and two
   * entries with one id would otherwise both appear.
   *
   * @param custom - Links loaded from storage
   */
  setCustomLinks(custom: SharePointLink[]): void {
    this.customLinkIds = new Set(custom.map((link) => link.id));
    this.links = [...this.baseLinks.filter((link) => !this.customLinkIds.has(link.id)), ...custom];
  }

  /**
   * Whether a link came from the user rather than the built-in registry
   */
  isCustomLink(id: string): boolean {
    return this.customLinkIds.has(id);
  }

  /**
   * Gets all links from the registry
   * @returns All registered links
   */
  getAllLinks(): SharePointLink[] {
    return [...this.links];
  }

  /**
   * Gets links filtered by category
   * @param category - Category to filter by
   * @returns Links in the specified category
   */
  getLinksByCategory(category: LinkCategory): SharePointLink[] {
    return this.links.filter((link) => link.category === category);
  }

  /**
   * Gets links that are applicable in the current context
   * @param values - Current placeholder values
   * @param isListContext - Whether user is in a list/library context
   * @param isSiteAdmin - Whether user has site admin permissions
   * @param isTenantAdmin - Whether user has tenant admin permissions
   * @param isModern - Whether site is modern SharePoint
   * @returns Applicable links
   */
  getApplicableLinks(
    values: PlaceholderValues,
    isListContext: boolean = false,
    isSiteAdmin: boolean = false,
    isTenantAdmin: boolean = false,
    isModern: boolean = true
  ): SharePointLink[] {
    console.log('LinkManager.getApplicableLinks called with:', {
      isListContext,
      hasListId: !!values.listId,
      hasListUrl: !!values.listUrl,
      placeholders: values,
    });

    const filtered = this.links.filter((link) => {
      // Check if link can be resolved with current values
      const canResolve = LinkTemplateEngine.canResolve(link.urlTemplate, values);
      if (!canResolve) {
        if (link.id === 'list-settings') {
          console.log('list-settings REJECTED: cannot resolve template', {
            template: link.urlTemplate,
            missingPlaceholders: LinkTemplateEngine.getMissingPlaceholders(
              link.urlTemplate,
              values
            ),
          });
        }
        return false;
      }

      // Check context requirements
      if (link.context) {
        if (link.context.requiresList && !isListContext) {
          if (link.id === 'list-settings') {
            console.log('list-settings REJECTED: requiresList=true but isListContext=false');
          }
          return false;
        }

        if (link.context.requiresSiteAdmin && !isSiteAdmin) {
          return false;
        }

        if (link.context.requiresTenantAdmin && !isTenantAdmin) {
          return false;
        }

        if (link.context.modernOnly && !isModern) {
          return false;
        }

        if (link.context.classicOnly && isModern) {
          return false;
        }
      }

      if (link.id === 'list-settings') {
        console.log('list-settings ACCEPTED!');
      }

      return true;
    });

    console.log('LinkManager.getApplicableLinks returning', filtered.length, 'links');
    return filtered;
  }

  /**
   * Resolves a link's URL template with placeholder values
   * @param link - Link to resolve
   * @param values - Placeholder values
   * @returns Resolved URL or null if cannot be resolved
   */
  resolveLink(link: SharePointLink, values: PlaceholderValues): string | null {
    return LinkTemplateEngine.resolve(link.urlTemplate, values);
  }

  /**
   * Groups links by category
   * @param links - Links to group
   * @returns Links grouped by category
   */
  groupByCategory(links: SharePointLink[]): GroupedLinks {
    const grouped: GroupedLinks = {};

    links.forEach((link) => {
      const category = link.category;
      if (!grouped[category]) {
        grouped[category] = [];
      }
      grouped[category].push(link);
    });

    return grouped;
  }

  /**
   * Sorts links by priority (lower number = higher priority)
   * @param links - Links to sort
   * @returns Sorted links
   */
  sortByPriority(links: SharePointLink[]): SharePointLink[] {
    return [...links].sort((a, b) => {
      const priorityA = a.priority ?? Number.MAX_SAFE_INTEGER;
      const priorityB = b.priority ?? Number.MAX_SAFE_INTEGER;
      return priorityA - priorityB;
    });
  }

  /**
   * Gets priority links (links with priority value set)
   * @returns Priority links sorted by priority
   */
  getPriorityLinks(): SharePointLink[] {
    const priorityLinks = this.links.filter((link) => link.priority !== undefined);
    return this.sortByPriority(priorityLinks);
  }

  /**
   * Searches links by title, description, or keywords
   * @param query - Search query
   * @param links - Links to search (defaults to all links)
   * @returns Matching links
   */
  search(query: string, links?: SharePointLink[]): SharePointLink[] {
    const searchLinks = links || this.links;
    const lowerQuery = query.toLowerCase();

    return searchLinks.filter((link) => {
      // Search in title
      if (link.title.toLowerCase().includes(lowerQuery)) {
        return true;
      }

      // Search in description
      if (link.description.toLowerCase().includes(lowerQuery)) {
        return true;
      }

      // Search in keywords
      if (link.keywords) {
        return link.keywords.some((keyword) => keyword.toLowerCase().includes(lowerQuery));
      }

      return false;
    });
  }

  /**
   * Gets a link by its ID
   * @param id - Link ID
   * @returns Link or undefined if not found
   */
  getLinkById(id: string): SharePointLink | undefined {
    return this.links.find((link) => link.id === id);
  }

  /**
   * Gets all available categories
   * @returns Array of all link categories
   */
  getCategories(): LinkCategory[] {
    return Object.values(LinkCategory);
  }

  /**
   * Track link access for recent links
   * @param link - Link that was accessed
   * @param resolvedUrl - The resolved URL
   */
  async trackLinkAccess(link: SharePointLink, resolvedUrl: string): Promise<void> {
    try {
      await LinkTracker.trackLinkAccess(link, resolvedUrl);
    } catch (error) {
      logError('Failed to track link access', error);
      // Don't throw - tracking failures shouldn't break navigation
    }
  }

  /**
   * Get recent links
   * @param limit - Maximum number of links to return
   */
  async getRecentLinks(limit?: number): Promise<any[]> {
    const result = await safely(
      () => LinkTracker.getRecentLinks(limit),
      'Failed to get recent links'
    );
    return result.success ? result.data : [];
  }

  /**
   * Get most accessed links
   * @param limit - Maximum number of links to return
   */
  async getMostAccessedLinks(limit?: number): Promise<any[]> {
    const result = await safely(
      () => LinkTracker.getMostAccessedLinks(limit),
      'Failed to get most accessed links'
    );
    return result.success ? result.data : [];
  }

  /**
   * Check if a link is favorited
   * @param linkId - Link ID
   */
  async isFavorite(linkId: string): Promise<boolean> {
    const result = await safely(
      () => FavoritesManager.isFavorite(linkId),
      'Failed to check favorite status'
    );
    return result.success ? result.data : false;
  }

  /**
   * Add a link to favorites
   * @param link - Link to favorite
   * @param note - Optional note
   */
  async addFavorite(link: SharePointLink, note?: string): Promise<void> {
    try {
      await FavoritesManager.addFavorite(link, note);
    } catch (error) {
      logError('Failed to add favorite', error);
      throw error; // Re-throw so UI can show error
    }
  }

  /**
   * Remove a link from favorites
   * @param linkId - Link ID
   */
  async removeFavorite(linkId: string): Promise<void> {
    try {
      await FavoritesManager.removeFavorite(linkId);
    } catch (error) {
      logError('Failed to remove favorite', error);
      throw error; // Re-throw so UI can show error
    }
  }

  /**
   * Toggle favorite status
   * @param link - Link to toggle
   * @returns New favorite status
   */
  async toggleFavorite(link: SharePointLink): Promise<boolean> {
    try {
      return await FavoritesManager.toggleFavorite(link);
    } catch (error) {
      logError('Failed to toggle favorite', error);
      throw error; // Re-throw so UI can show error
    }
  }

  /**
   * Get all favorite links
   */
  async getFavorites(): Promise<any[]> {
    const result = await safely(() => FavoritesManager.getFavorites(), 'Failed to get favorites');
    return result.success ? result.data : [];
  }

  /**
   * Checks if a link meets context requirements
   * @param link - Link to check
   * @param context - Current context
   * @returns True if link is applicable
   */
  private meetsContextRequirements(
    link: SharePointLink,
    context: {
      isListContext: boolean;
      isSiteAdmin: boolean;
      isTenantAdmin: boolean;
      isModern: boolean;
    }
  ): boolean {
    if (!link.context) {
      return true;
    }

    const { isListContext, isSiteAdmin, isTenantAdmin, isModern } = context;

    if (link.context.requiresList && !isListContext) {
      return false;
    }

    if (link.context.requiresSiteAdmin && !isSiteAdmin) {
      return false;
    }

    if (link.context.requiresTenantAdmin && !isTenantAdmin) {
      return false;
    }

    if (link.context.modernOnly && !isModern) {
      return false;
    }

    if (link.context.classicOnly && isModern) {
      return false;
    }

    return true;
  }
}

// Export a singleton instance for convenience
export const linkManager = new LinkManager();
