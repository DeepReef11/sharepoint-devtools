/**
 * QuickNav Modal Component
 * Implements keyboard-driven navigation modal with fuzzy search
 * This is a wrapper around the shared Modal UI component from src/ui/modal
 */

import { Modal, ResultItem, CategoryGroup } from '../ui/modal';

export interface NavigationLink {
  id: string;
  title: string;
  category: string;
  url: string;
  description?: string;
  applicable: boolean;
}

export class QuickNavModal {
  private modal: Modal;
  private links: NavigationLink[] = [];
  private filteredLinks: NavigationLink[] = [];
  private categoryShortcuts: Map<string, string> = new Map();
  private currentQuery: string = '';

  constructor() {
    // Initialize the shared modal component
    this.modal = new Modal({
      title: 'SharePoint DevTools',
      placeholder: 'Search SharePoint locations... (or use arrow keys)',
      onSearch: (query: string) => this.handleSearch(query),
      onSelect: (item: ResultItem) => this.handleSelect(item),
      onClose: () => this.handleClose(),
    });

    this.setupKeyboardShortcuts();
    this.initializeSampleData();
  }

  /**
   * Initialize sample navigation links for testing
   * TODO: Replace with actual SharePoint links registry in Issue #4
   */
  private initializeSampleData(): void {
    this.links = [
      {
        id: 'site-settings',
        title: 'Site Settings',
        category: 'Site Admin',
        url: '_layouts/15/settings.aspx',
        description: 'Access site settings and configuration',
        applicable: true,
      },
      {
        id: 'site-contents',
        title: 'Site Contents',
        category: 'Content',
        url: '_layouts/15/viewlsts.aspx',
        description: 'View all lists, libraries, and apps',
        applicable: true,
      },
      {
        id: 'site-permissions',
        title: 'Site Permissions',
        category: 'Site Admin',
        url: '_layouts/15/user.aspx',
        description: 'Manage site permissions and users',
        applicable: true,
      },
      {
        id: 'recycle-bin',
        title: 'Recycle Bin',
        category: 'Content',
        url: '_layouts/15/RecycleBin.aspx',
        description: 'View and restore deleted items',
        applicable: true,
      },
      {
        id: 'pages-library',
        title: 'Pages Library',
        category: 'Design',
        url: 'SitePages',
        description: 'Manage site pages',
        applicable: true,
      },
      {
        id: 'site-navigation',
        title: 'Navigation Settings',
        category: 'Navigation',
        url: '_layouts/15/AreaNavigationSettings.aspx',
        description: 'Configure site navigation',
        applicable: true,
      },
      {
        id: 'regional-settings',
        title: 'Regional Settings',
        category: 'Site Admin',
        url: '_layouts/15/regionalsetng.aspx',
        description: 'Configure time zone, locale, and calendar',
        applicable: true,
      },
      {
        id: 'site-features',
        title: 'Site Features',
        category: 'Site Admin',
        url: '_layouts/15/ManageFeatures.aspx',
        description: 'Activate or deactivate site features',
        applicable: true,
      },
      {
        id: 'search-settings',
        title: 'Search Settings',
        category: 'Advanced',
        url: '_layouts/15/enhancedSearch.aspx',
        description: 'Configure search settings',
        applicable: true,
      },
    ];

    // Setup category shortcuts (Ctrl+1, Ctrl+2, etc.)
    const categories = [...new Set(this.links.map((link) => link.category))];
    categories.forEach((category, index) => {
      if (index < 9) {
        this.categoryShortcuts.set(`${index + 1}`, category);
      }
    });

    this.filteredLinks = [...this.links];
    this.updateModalResults();
  }

  /**
   * Setup keyboard shortcut to toggle modal (Ctrl+K / Cmd+K)
   */
  private setupKeyboardShortcuts(): void {
    document.addEventListener('keydown', (e: KeyboardEvent) => {
      // Toggle modal with Ctrl+K or Cmd+K
      if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
        e.preventDefault();
        this.toggle();
        return;
      }

      // Category shortcuts (Ctrl+1 through Ctrl+9)
      if (this.isOpen() && (e.ctrlKey || e.metaKey) && /^[1-9]$/.test(e.key)) {
        e.preventDefault();
        this.filterByCategory(e.key);
        return;
      }
    });
  }

  /**
   * Convert NavigationLink to ResultItem
   */
  private linkToResultItem(link: NavigationLink): ResultItem {
    return {
      id: link.id,
      title: link.title,
      description: link.description,
      url: link.url,
      category: link.category,
    };
  }

  /**
   * Update modal with current filtered results
   */
  private updateModalResults(): void {
    // Group by category
    const grouped = this.groupByCategory(this.filteredLinks);
    const categoryGroups: CategoryGroup[] = Object.entries(grouped).map(([category, links]) => ({
      category,
      items: links.map((link) => this.linkToResultItem(link)),
    }));

    this.modal.setGroupedResults(categoryGroups);
  }

  /**
   * Handle search input
   * TODO: Implement fuzzy search with Fuse.js in Issue #8
   */
  private handleSearch(query: string): void {
    this.currentQuery = query;

    if (!query.trim()) {
      this.filteredLinks = [...this.links];
    } else {
      // Simple substring search for now
      const lowerQuery = query.toLowerCase();
      this.filteredLinks = this.links.filter(
        (link) =>
          link.title.toLowerCase().includes(lowerQuery) ||
          link.category.toLowerCase().includes(lowerQuery) ||
          link.description?.toLowerCase().includes(lowerQuery)
      );
    }

    this.updateModalResults();
  }

  /**
   * Handle item selection
   */
  private handleSelect(item: ResultItem): void {
    // Build full URL
    const baseUrl =
      window.location.origin + window.location.pathname.split('/').slice(0, -1).join('/');
    const fullUrl = item.url?.startsWith('http') ? item.url : `${baseUrl}/${item.url}`;

    window.location.href = fullUrl;
  }

  /**
   * Handle modal close
   */
  private handleClose(): void {
    this.currentQuery = '';
    this.filteredLinks = [...this.links];
  }

  /**
   * Filter results by category (using Ctrl+1-9)
   */
  private filterByCategory(key: string): void {
    const category = this.categoryShortcuts.get(key);
    if (!category) return;

    this.filteredLinks = this.links.filter((link) => link.category === category);
    this.updateModalResults();
  }

  /**
   * Group links by category
   */
  private groupByCategory(links: NavigationLink[]): Record<string, NavigationLink[]> {
    return links.reduce(
      (acc, link) => {
        if (!acc[link.category]) {
          acc[link.category] = [];
        }
        acc[link.category].push(link);
        return acc;
      },
      {} as Record<string, NavigationLink[]>
    );
  }

  /**
   * Show the modal
   */
  public show(): void {
    this.modal.show();
  }

  /**
   * Hide the modal
   */
  public hide(): void {
    this.modal.close();
  }

  /**
   * Toggle modal visibility
   */
  public toggle(): void {
    this.modal.toggle();
  }

  /**
   * Check if modal is visible
   */
  public isOpen(): boolean {
    return this.modal.isOpen();
  }
}
