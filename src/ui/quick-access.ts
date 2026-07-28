/**
 * Quick Access UI Component
 * Displays recent and favorite links in a quick access section
 */

import { LinkTracker, FavoritesManager } from '../storage';
import { escapeHtml as escapeHtmlValue } from '../utils/html';
import { linkManager } from '../links/link-manager';
import { PlaceholderValues } from '../links/types';

export interface QuickAccessOptions {
  onLinkClick?: (linkId: string, url: string) => void;
  placeholderValues?: PlaceholderValues;
}

/**
 * Quick Access Component for Recent and Favorite Links
 */
export class QuickAccess {
  private container: HTMLElement | null = null;
  private options: QuickAccessOptions;

  constructor(options: QuickAccessOptions = {}) {
    this.options = options;
  }

  /**
   * Render the quick access section
   */
  async render(containerElement: HTMLElement): Promise<void> {
    this.container = containerElement;

    const [recentLinks, favoriteLinks] = await Promise.all([
      LinkTracker.getRecentLinks(5),
      FavoritesManager.getFavorites(),
    ]);

    const html = `
      <div class="sp-quick-access">
        ${favoriteLinks.length > 0 ? this.renderFavorites(favoriteLinks.slice(0, 5)) : ''}
        ${recentLinks.length > 0 ? this.renderRecent(recentLinks) : ''}
        ${favoriteLinks.length === 0 && recentLinks.length === 0 ? this.renderEmpty() : ''}
      </div>
    `;

    containerElement.innerHTML = html;
    this.attachEventListeners();
  }

  /**
   * Render favorites section
   */
  private renderFavorites(favorites: any[]): string {
    return `
      <div class="sp-qa-section sp-qa-favorites">
        <h4 class="sp-qa-heading">
          ⭐ Favorites
        </h4>
        <ul class="sp-qa-list">
          ${favorites.map((fav) => this.renderFavoriteItem(fav)).join('')}
        </ul>
      </div>
    `;
  }

  /**
   * Render a favorite link item
   */
  private renderFavoriteItem(favorite: any): string {
    return `
      <li class="sp-qa-item sp-qa-favorite-item" data-link-id="${this.escapeHtml(favorite.linkId)}">
        <button class="sp-qa-link" data-link-id="${this.escapeHtml(favorite.linkId)}">
          <span class="sp-qa-title">${this.escapeHtml(favorite.title)}</span>
          <span class="sp-qa-category">${this.escapeHtml(favorite.category)}</span>
        </button>
        <button class="sp-qa-unfavorite" data-link-id="${this.escapeHtml(favorite.linkId)}" title="Remove from favorites">
          ★
        </button>
      </li>
    `;
  }

  /**
   * Render recent links section
   */
  private renderRecent(recentLinks: any[]): string {
    return `
      <div class="sp-qa-section sp-qa-recent">
        <h4 class="sp-qa-heading">
          🕒 Recent
        </h4>
        <ul class="sp-qa-list">
          ${recentLinks.map((link) => this.renderRecentItem(link)).join('')}
        </ul>
      </div>
    `;
  }

  /**
   * Render a recent link item
   */
  private renderRecentItem(link: any): string {
    return `
      <li class="sp-qa-item sp-qa-recent-item" data-link-id="${this.escapeHtml(link.linkId)}">
        <button class="sp-qa-link" data-link-id="${this.escapeHtml(link.linkId)}">
          <span class="sp-qa-title">${this.escapeHtml(link.title)}</span>
          <span class="sp-qa-meta">
            <span class="sp-qa-category">${this.escapeHtml(link.category)}</span>
            ${link.accessCount > 1 ? `<span class="sp-qa-count">${link.accessCount}×</span>` : ''}
          </span>
        </button>
        <button class="sp-qa-favorite-btn" data-link-id="${this.escapeHtml(link.linkId)}" title="Add to favorites">
          ☆
        </button>
      </li>
    `;
  }

  /**
   * Render empty state
   */
  private renderEmpty(): string {
    return `
      <div class="sp-qa-empty">
        <p>No recent or favorite links yet.</p>
        <p class="sp-qa-empty-hint">Links you access will appear here for quick access.</p>
      </div>
    `;
  }

  /**
   * Attach event listeners
   */
  private attachEventListeners(): void {
    if (!this.container) return;

    // Link click handlers
    this.container.querySelectorAll('.sp-qa-link').forEach((btn) => {
      btn.addEventListener('click', async (e) => {
        e.preventDefault();
        const linkId = (btn as HTMLElement).dataset.linkId;
        if (linkId) {
          await this.handleLinkClick(linkId);
        }
      });
    });

    // Unfavorite handlers
    this.container.querySelectorAll('.sp-qa-unfavorite').forEach((btn) => {
      btn.addEventListener('click', async (e) => {
        e.preventDefault();
        e.stopPropagation();
        const linkId = (btn as HTMLElement).dataset.linkId;
        if (linkId) {
          await this.handleUnfavorite(linkId);
        }
      });
    });

    // Favorite handlers
    this.container.querySelectorAll('.sp-qa-favorite-btn').forEach((btn) => {
      btn.addEventListener('click', async (e) => {
        e.preventDefault();
        e.stopPropagation();
        const linkId = (btn as HTMLElement).dataset.linkId;
        if (linkId) {
          await this.handleFavorite(linkId);
        }
      });
    });
  }

  /**
   * Handle link click
   */
  private async handleLinkClick(linkId: string): Promise<void> {
    const link = linkManager.getLinkById(linkId);
    if (!link) return;

    const url = this.options.placeholderValues
      ? linkManager.resolveLink(link, this.options.placeholderValues)
      : null;

    if (url) {
      if (this.options.onLinkClick) {
        this.options.onLinkClick(linkId, url);
      } else {
        // Default behavior: open in new tab
        window.open(url, '_blank');
      }

      // Track the link access
      await LinkTracker.trackLinkAccess(link, url);
    }
  }

  /**
   * Handle unfavorite
   */
  private async handleUnfavorite(linkId: string): Promise<void> {
    await FavoritesManager.removeFavorite(linkId);
    // Re-render to update UI
    if (this.container) {
      await this.render(this.container);
    }
  }

  /**
   * Handle favorite
   */
  private async handleFavorite(linkId: string): Promise<void> {
    const link = linkManager.getLinkById(linkId);
    if (link) {
      await FavoritesManager.addFavorite(link);
      // Re-render to update UI
      if (this.container) {
        await this.render(this.container);
      }
    }
  }

  /**
   * Escape HTML to prevent XSS
   */
  private escapeHtml(text: string): string {
    return escapeHtmlValue(text);
  }

  /**
   * Inject styles for quick access
   */
  static injectStyles(): void {
    if (document.getElementById('sp-quick-access-styles')) {
      return;
    }

    const style = document.createElement('style');
    style.id = 'sp-quick-access-styles';
    style.textContent = `
      .sp-quick-access {
        padding: 16px 0;
      }

      .sp-qa-section {
        margin-bottom: 24px;
      }

      .sp-qa-section:last-child {
        margin-bottom: 0;
      }

      .sp-qa-heading {
        margin: 0 0 12px 0;
        font-size: 14px;
        font-weight: 600;
        color: #333;
        display: flex;
        align-items: center;
        gap: 6px;
      }

      .sp-qa-list {
        list-style: none;
        margin: 0;
        padding: 0;
      }

      .sp-qa-item {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 4px;
      }

      .sp-qa-link {
        flex: 1;
        display: flex;
        flex-direction: column;
        align-items: flex-start;
        padding: 8px 12px;
        background: #f5f5f5;
        border: 1px solid #e0e0e0;
        border-radius: 6px;
        cursor: pointer;
        text-align: left;
        transition: all 0.2s;
      }

      .sp-qa-link:hover {
        background: #e8e8e8;
        border-color: #0078d4;
      }

      .sp-qa-title {
        font-size: 13px;
        font-weight: 500;
        color: #0078d4;
        margin-bottom: 2px;
      }

      .sp-qa-meta {
        display: flex;
        gap: 8px;
        align-items: center;
      }

      .sp-qa-category {
        font-size: 11px;
        color: #666;
      }

      .sp-qa-count {
        font-size: 10px;
        color: #999;
        background: #fff;
        padding: 2px 6px;
        border-radius: 10px;
      }

      .sp-qa-unfavorite,
      .sp-qa-favorite-btn {
        padding: 4px 8px;
        background: transparent;
        border: none;
        font-size: 18px;
        cursor: pointer;
        opacity: 0.6;
        transition: opacity 0.2s;
      }

      .sp-qa-unfavorite {
        color: #ffc107;
      }

      .sp-qa-favorite-btn {
        color: #999;
      }

      .sp-qa-unfavorite:hover,
      .sp-qa-favorite-btn:hover {
        opacity: 1;
      }

      .sp-qa-empty {
        text-align: center;
        padding: 32px 16px;
        color: #666;
      }

      .sp-qa-empty p {
        margin: 0 0 8px 0;
      }

      .sp-qa-empty-hint {
        font-size: 12px;
        color: #999;
      }

      @media (prefers-color-scheme: dark) {
        .sp-qa-heading {
          color: #fff;
        }

        .sp-qa-link {
          background: #2d2d2d;
          border-color: #444;
        }

        .sp-qa-link:hover {
          background: #3d3d3d;
          border-color: #0078d4;
        }

        .sp-qa-title {
          color: #4db8ff;
        }

        .sp-qa-category {
          color: #aaa;
        }

        .sp-qa-count {
          background: #1a1a1a;
          color: #999;
        }

        .sp-qa-empty {
          color: #aaa;
        }

        .sp-qa-empty-hint {
          color: #777;
        }
      }
    `;

    document.head.appendChild(style);
  }
}
