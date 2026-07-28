/**
 * Column Inspector Integration for Content Script
 * Initializes and manages the column inspector
 */

import { ColumnInspector } from '../ui/column-inspector';
import { ColumnFetcher } from '../api/column-fetcher';

export class ColumnInspectorIntegration {
  private inspector: ColumnInspector | null = null;
  private floatingButton: HTMLElement | null = null;

  /**
   * Initialize the column inspector
   */
  init(): void {
    // Only initialize on SharePoint pages
    if (!this.isSharePointPage()) {
      return;
    }

    // Create the inspector
    this.inspector = new ColumnInspector();

    // Listen for keyboard shortcut (Alt+C) - supports both left and right Alt
    document.addEventListener('keydown', (e) => {
      if ((e.altKey || e.getModifierState('AltGraph')) && e.key.toLowerCase() === 'c') {
        e.preventDefault();
        this.toggleInspector();
      }
    });
  }

  /**
   * Check if we're on a SharePoint page
   */
  private isSharePointPage(): boolean {
    const url = window.location.href;
    return url.includes('.sharepoint.com') || url.includes('.sharepoint-df.com');
  }

  /**
   * Create floating button to toggle inspector
   */
  private createFloatingButton(): void {
    const button = document.createElement('button');
    button.id = 'sp-inspector-toggle';
    button.className = 'sp-inspector-toggle';
    button.title = 'Toggle Column Inspector (Alt+C)';
    button.innerHTML = '📋';
    button.style.display = 'none';

    button.addEventListener('click', () => {
      this.toggleInspector();
    });

    document.body.appendChild(button);
    this.floatingButton = button;

    // Add styles
    this.injectFloatingButtonStyles();
  }

  /**
   * Inject styles for floating button
   */
  private injectFloatingButtonStyles(): void {
    if (document.getElementById('sp-inspector-toggle-styles')) {
      return;
    }

    const style = document.createElement('style');
    style.id = 'sp-inspector-toggle-styles';
    style.textContent = `
      .sp-inspector-toggle {
        position: fixed;
        bottom: 20px;
        right: 20px;
        width: 50px;
        height: 50px;
        border-radius: 50%;
        background: #0078d4;
        color: white;
        border: none;
        font-size: 24px;
        cursor: pointer;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
        z-index: 9999;
        transition: all 0.3s;
        display: flex;
        align-items: center;
        justify-content: center;
      }

      .sp-inspector-toggle:hover {
        transform: scale(1.1);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
      }

      .sp-inspector-toggle:active {
        transform: scale(0.95);
      }

      @media (prefers-color-scheme: dark) {
        .sp-inspector-toggle {
          background: #106ebe;
        }
      }
    `;

    document.head.appendChild(style);
  }

  /**
   * Toggle the inspector
   */
  private toggleInspector(): void {
    if (!this.inspector) return;
    this.inspector.toggle();
  }

  /**
   * Update floating button visibility based on context
   */
  private updateFloatingButton(isListView: boolean): void {
    if (!this.floatingButton) return;

    if (isListView) {
      this.floatingButton.style.display = 'flex';
    } else {
      this.floatingButton.style.display = 'none';
      // Hide inspector if it's showing
      if (this.inspector) {
        this.inspector.hide();
      }
    }
  }

  /**
   * Monitor URL changes in SharePoint SPA
   */
  private monitorUrlChanges(): void {
    let lastUrl = window.location.href;

    // Use MutationObserver to detect URL changes
    const observer = new MutationObserver(() => {
      const currentUrl = window.location.href;
      if (currentUrl !== lastUrl) {
        lastUrl = currentUrl;
        this.handleUrlChange();
      }
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true,
    });

    // Also listen to popstate (back/forward)
    window.addEventListener('popstate', () => {
      this.handleUrlChange();
    });

    // Listen to pushState and replaceState
    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    history.pushState = function (...args) {
      originalPushState.apply(this, args);
      window.dispatchEvent(new Event('pushstate'));
    };

    history.replaceState = function (...args) {
      originalReplaceState.apply(this, args);
      window.dispatchEvent(new Event('replacestate'));
    };

    window.addEventListener('pushstate', () => {
      this.handleUrlChange();
    });

    window.addEventListener('replacestate', () => {
      this.handleUrlChange();
    });
  }

  /**
   * Handle URL changes
   */
  private handleUrlChange(): void {
    // Check if we're on a list view
    const isListView = ColumnFetcher.isListView();
    this.updateFloatingButton(isListView);
  }
}
