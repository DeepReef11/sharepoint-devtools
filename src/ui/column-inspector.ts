/**
 * Column Inspector UI Component
 * Displays column metadata in a side panel
 */

import { ColumnSchema, ListSchema } from '../types/column-schema';
import { escapeHtml as escapeHtmlValue, sanitizeUrl } from '../utils/html';
import { ColumnFetcher } from '../api/column-fetcher';
import { extractListIdAsync } from '../context/sharepoint-context';

export class ColumnInspector {
  private container: HTMLElement | null = null;
  private visible: boolean = false;
  private currentList: ListSchema | null = null;
  private currentItem: Record<string, any> | null = null;
  private currentItemIdOrPath: number | string | null = null;
  private filterText: string = '';
  private showHidden: boolean = false;
  private selectedColumnIndex: number = -1;
  private innerNavigationMode: boolean = false;
  private innerFocusIndex: number = -1;

  constructor() {
    this.createUI();
    this.attachEventListeners();
  }

  /**
   * Create the inspector UI
   */
  private createUI(): void {
    const inspector = document.createElement('div');
    inspector.id = 'sp-column-inspector';
    inspector.className = 'sp-column-inspector';
    inspector.style.display = 'none';

    inspector.innerHTML = `
      <div class="sp-inspector-header">
        <div class="sp-inspector-title-section">
          <h3 id="sp-inspector-title">Column Inspector</h3>
          <div id="sp-inspector-list-name" class="sp-inspector-list-name" style="display: none;"></div>
        </div>
        <div class="sp-inspector-controls">
          <button id="sp-inspector-refresh" class="sp-btn" title="Refresh">
            🔄
          </button>
          <button id="sp-inspector-close" class="sp-btn" title="Close">
            ✕
          </button>
        </div>
      </div>
      <div class="sp-inspector-toolbar">
        <input
          type="text"
          id="sp-inspector-search"
          class="sp-search-input"
          placeholder="Filter columns..."
        />
        <label class="sp-checkbox-label">
          <input type="checkbox" id="sp-inspector-show-hidden" />
          Show hidden
        </label>
      </div>
      <div class="sp-inspector-content">
        <div id="sp-inspector-loading" class="sp-loading">
          Loading column schema...
        </div>
        <div id="sp-inspector-error" class="sp-error" style="display: none;">
        </div>
        <div id="sp-inspector-list-info" class="sp-list-info" style="display: none;">
        </div>
        <div id="sp-inspector-columns" class="sp-columns-grid" style="display: none;">
        </div>
      </div>
    `;

    this.container = inspector;
    document.body.appendChild(inspector);

    // Add styles
    this.injectStyles();
  }

  /**
   * Inject CSS styles
   */
  private injectStyles(): void {
    if (document.getElementById('sp-inspector-styles')) {
      return;
    }

    const style = document.createElement('style');
    style.id = 'sp-inspector-styles';
    style.textContent = `
      .sp-column-inspector {
        position: fixed;
        top: 0;
        right: 0;
        width: 500px;
        height: 100vh;
        background: #ffffff;
        box-shadow: -2px 0 10px rgba(0, 0, 0, 0.1);
        z-index: 10000;
        display: flex;
        flex-direction: column;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
        font-size: 14px;
      }

      @media (prefers-color-scheme: dark) {
        .sp-column-inspector {
          background: #1e1e1e;
          color: #d4d4d4;
          box-shadow: -2px 0 10px rgba(0, 0, 0, 0.5);
        }
      }

      .sp-inspector-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 16px;
        border-bottom: 1px solid #e0e0e0;
      }

      @media (prefers-color-scheme: dark) {
        .sp-inspector-header {
          border-bottom-color: #333;
        }
      }

      .sp-inspector-title-section {
        flex: 1;
        min-width: 0;
      }

      .sp-inspector-header h3 {
        margin: 0;
        font-size: 18px;
        font-weight: 600;
      }

      .sp-inspector-list-name {
        margin-top: 6px;
        font-size: 13px;
        color: #666;
        display: flex;
        align-items: center;
        gap: 6px;
      }

      @media (prefers-color-scheme: dark) {
        .sp-inspector-list-name {
          color: #999;
        }
      }

      .sp-list-internal-label {
        font-weight: 500;
      }

      .sp-list-internal-value {
        font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
        background: #f5f5f5;
        padding: 2px 6px;
        border-radius: 3px;
      }

      @media (prefers-color-scheme: dark) {
        .sp-list-internal-value {
          background: #2a2a2a;
        }
      }

      .sp-copy-btn-inline {
        background: transparent;
        border: 1px solid #ccc;
        border-radius: 3px;
        padding: 2px 8px;
        cursor: pointer;
        font-size: 11px;
        transition: background 0.2s;
        margin-left: 4px;
      }

      .sp-copy-btn-inline:hover {
        background: #e8e8e8;
      }

      @media (prefers-color-scheme: dark) {
        .sp-copy-btn-inline {
          border-color: #555;
          color: #d4d4d4;
        }
        .sp-copy-btn-inline:hover {
          background: #333;
        }
      }

      .sp-inspector-controls {
        display: flex;
        gap: 8px;
      }

      .sp-btn {
        background: transparent;
        border: 1px solid #ccc;
        border-radius: 4px;
        padding: 6px 12px;
        cursor: pointer;
        font-size: 14px;
        transition: background 0.2s;
      }

      .sp-btn:hover {
        background: #f5f5f5;
      }

      @media (prefers-color-scheme: dark) {
        .sp-btn {
          border-color: #555;
          color: #d4d4d4;
        }
        .sp-btn:hover {
          background: #2a2a2a;
        }
      }

      .sp-inspector-toolbar {
        padding: 12px 16px;
        border-bottom: 1px solid #e0e0e0;
        display: flex;
        gap: 12px;
        align-items: center;
      }

      @media (prefers-color-scheme: dark) {
        .sp-inspector-toolbar {
          border-bottom-color: #333;
        }
      }

      .sp-search-input {
        flex: 1;
        padding: 8px 12px;
        border: 1px solid #ccc;
        border-radius: 4px;
        font-size: 14px;
      }

      @media (prefers-color-scheme: dark) {
        .sp-search-input {
          background: #2a2a2a;
          border-color: #555;
          color: #d4d4d4;
        }
      }

      .sp-checkbox-label {
        display: flex;
        align-items: center;
        gap: 6px;
        white-space: nowrap;
        cursor: pointer;
      }

      .sp-inspector-content {
        flex: 1;
        overflow-y: auto;
        padding: 16px;
      }

      .sp-loading {
        text-align: center;
        padding: 40px 20px;
        color: #666;
      }

      @media (prefers-color-scheme: dark) {
        .sp-loading {
          color: #999;
        }
      }

      .sp-error {
        padding: 16px;
        background: #fee;
        border: 1px solid #fcc;
        border-radius: 4px;
        color: #c00;
      }

      @media (prefers-color-scheme: dark) {
        .sp-error {
          background: #4a1515;
          border-color: #6a2020;
          color: #ff6b6b;
        }
      }

      .sp-list-info {
        margin-bottom: 20px;
        padding: 16px;
        background: #f9f9f9;
        border-radius: 4px;
      }

      @media (prefers-color-scheme: dark) {
        .sp-list-info {
          background: #2a2a2a;
        }
      }

      .sp-list-info h4 {
        margin: 0 0 12px 0;
        font-size: 16px;
      }

      .sp-management-links {
        display: flex;
        gap: 8px;
        margin: 12px 0;
        flex-wrap: wrap;
      }

      .sp-link-btn {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 8px 16px;
        background: #0078d4;
        color: white;
        text-decoration: none;
        border-radius: 4px;
        font-size: 13px;
        font-weight: 500;
        transition: background 0.2s;
      }

      .sp-link-btn:hover {
        background: #106ebe;
      }

      @media (prefers-color-scheme: dark) {
        .sp-link-btn {
          background: #106ebe;
        }
        .sp-link-btn:hover {
          background: #005a9e;
        }
      }

      .sp-content-types-list {
        flex-direction: column;
        gap: 8px;
      }

      .sp-content-types-links {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
      }

      .sp-content-type-link {
        display: inline-flex;
        align-items: center;
        padding: 4px 10px;
        background: #f0f0f0;
        color: #333;
        text-decoration: none;
        border-radius: 3px;
        font-size: 12px;
        transition: background 0.2s;
        border: 1px solid #d0d0d0;
      }

      .sp-content-type-link:hover {
        background: #e0e0e0;
        border-color: #b0b0b0;
      }

      @media (prefers-color-scheme: dark) {
        .sp-content-type-link {
          background: #2a2a2a;
          color: #d4d4d4;
          border-color: #555;
        }
        .sp-content-type-link:hover {
          background: #3a3a3a;
          border-color: #666;
        }
      }

      .sp-info-row {
        display: flex;
        margin: 6px 0;
      }

      .sp-info-label {
        font-weight: 600;
        width: 140px;
        color: #666;
      }

      @media (prefers-color-scheme: dark) {
        .sp-info-label {
          color: #999;
        }
      }

      .sp-info-value {
        flex: 1;
        font-family: 'Courier New', monospace;
        font-size: 13px;
      }

      .sp-columns-grid {
        display: flex;
        flex-direction: column;
        gap: 12px;
      }

      .sp-column-card {
        border: 1px solid #e0e0e0;
        border-radius: 6px;
        padding: 12px;
        background: #fafafa;
        cursor: pointer;
        transition: background 0.2s, border-color 0.2s, box-shadow 0.2s;
      }

      @media (prefers-color-scheme: dark) {
        .sp-column-card {
          background: #2a2a2a;
          border-color: #444;
        }
      }

      .sp-column-card:hover {
        background: #f0f0f0;
        border-color: #ccc;
      }

      @media (prefers-color-scheme: dark) {
        .sp-column-card:hover {
          background: #333;
          border-color: #555;
        }
      }

      .sp-column-card.selected {
        background: #e3f2fd;
        border-color: #1976d2;
        border-width: 2px;
        box-shadow: 0 0 0 2px rgba(25, 118, 210, 0.2);
      }

      @media (prefers-color-scheme: dark) {
        .sp-column-card.selected {
          background: #1a3a52;
          border-color: #64b5f6;
          box-shadow: 0 0 0 2px rgba(100, 181, 246, 0.2);
        }
      }

      .sp-column-card.inner-nav-active {
        background: #bbdefb;
        border-color: #0d47a1;
        border-width: 2px;
        box-shadow: 0 0 0 3px rgba(13, 71, 161, 0.3);
      }

      @media (prefers-color-scheme: dark) {
        .sp-column-card.inner-nav-active {
          background: #1a4a6a;
          border-color: #90caf9;
          box-shadow: 0 0 0 3px rgba(144, 202, 249, 0.3);
        }
      }

      .sp-column-card.inner-nav-active .sp-copy-btn:focus,
      .sp-column-card.inner-nav-active .sp-edit-column-link:focus {
        outline: 2px solid #0d47a1;
        outline-offset: 2px;
        background: #e3f2fd;
      }

      @media (prefers-color-scheme: dark) {
        .sp-column-card.inner-nav-active .sp-copy-btn:focus,
        .sp-column-card.inner-nav-active .sp-edit-column-link:focus {
          outline-color: #90caf9;
          background: #1a3a52;
        }
      }

      .sp-column-card.hidden {
        opacity: 0.6;
        border-style: dashed;
      }

      .sp-column-header {
        display: flex;
        justify-content: space-between;
        align-items: start;
        margin-bottom: 10px;
      }

      .sp-column-title {
        font-weight: 600;
        font-size: 15px;
      }

      .sp-column-actions {
        margin: 10px 0;
      }

      .sp-edit-column-link {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 12px;
        background: #f3f2f1;
        color: #323130;
        text-decoration: none;
        border-radius: 3px;
        font-size: 12px;
        font-weight: 500;
        transition: background 0.2s;
        border: 1px solid #e1dfdd;
      }

      .sp-edit-column-link:hover {
        background: #edebe9;
        border-color: #d2d0ce;
      }

      @media (prefers-color-scheme: dark) {
        .sp-edit-column-link {
          background: #3a3a3a;
          color: #d4d4d4;
          border-color: #555;
        }
        .sp-edit-column-link:hover {
          background: #4a4a4a;
          border-color: #666;
        }
      }

      .sp-column-badges {
        display: flex;
        gap: 4px;
        flex-wrap: wrap;
      }

      .sp-badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 500;
        text-transform: uppercase;
      }

      .sp-badge.required {
        background: #e3f2fd;
        color: #1976d2;
      }

      .sp-badge.hidden {
        background: #fafafa;
        color: #666;
        border: 1px solid #ddd;
      }

      .sp-badge.readonly {
        background: #fff3e0;
        color: #f57c00;
      }

      .sp-badge.indexed {
        background: #f3e5f5;
        color: #7b1fa2;
      }

      @media (prefers-color-scheme: dark) {
        .sp-badge.required {
          background: #1a3a52;
          color: #64b5f6;
        }
        .sp-badge.hidden {
          background: #2a2a2a;
          color: #999;
          border-color: #444;
        }
        .sp-badge.readonly {
          background: #3a2a1a;
          color: #ffb74d;
        }
        .sp-badge.indexed {
          background: #2a1a3a;
          color: #ba68c8;
        }
      }

      .sp-column-details {
        display: grid;
        grid-template-columns: auto 1fr;
        gap: 8px 12px;
        font-size: 13px;
      }

      .sp-current-value-row {
        display: flex;
        flex-direction: column;
        gap: 6px;
        grid-column: 1 / -1;
        margin-top: 8px;
        padding-top: 8px;
        border-top: 2px solid #0078d4;
      }

      @media (prefers-color-scheme: dark) {
        .sp-current-value-row {
          border-top-color: #106ebe;
        }
      }

      .sp-current-value-label {
        color: #0078d4;
        font-weight: 600;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
      }

      @media (prefers-color-scheme: dark) {
        .sp-current-value-label {
          color: #64b5f6;
        }
      }

      .sp-current-value-content {
        background: #f3f2f1;
        padding: 10px 12px;
        border-radius: 4px;
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
        font-size: 14px;
        line-height: 1.6;
        word-wrap: break-word;
        overflow-wrap: break-word;
        white-space: pre-wrap;
        max-width: 100%;
        border: 1px solid #e1e1e1;
      }

      @media (prefers-color-scheme: dark) {
        .sp-current-value-content {
          background: #2a2a2a;
          border-color: #444;
        }
      }

      .sp-current-value-content a {
        color: #0078d4;
        text-decoration: underline;
      }

      .sp-current-value-content a:hover {
        color: #106ebe;
      }

      @media (prefers-color-scheme: dark) {
        .sp-current-value-content a {
          color: #64b5f6;
        }
        .sp-current-value-content a:hover {
          color: #90caf9;
        }
      }

      .sp-current-value-actions {
        display: flex;
        gap: 8px;
        margin-top: 4px;
      }

      .sp-value-copy-btn {
        padding: 6px 12px;
        background: #0078d4;
        color: white;
        border: none;
        border-radius: 3px;
        cursor: pointer;
        font-size: 12px;
        font-weight: 500;
        transition: background 0.2s;
      }

      .sp-value-copy-btn:hover {
        background: #106ebe;
      }

      @media (prefers-color-scheme: dark) {
        .sp-value-copy-btn {
          background: #106ebe;
        }
        .sp-value-copy-btn:hover {
          background: #005a9e;
        }
      }

      .sp-detail-label {
        color: #666;
        font-weight: 500;
      }

      @media (prefers-color-scheme: dark) {
        .sp-detail-label {
          color: #999;
        }
      }

      .sp-detail-value {
        font-family: 'Courier New', monospace;
        word-break: break-all;
      }

      .sp-copy-btn {
        background: none;
        border: none;
        color: #1976d2;
        cursor: pointer;
        padding: 0;
        margin-left: 8px;
        font-size: 12px;
      }

      .sp-copy-btn:hover {
        text-decoration: underline;
      }

      @media (prefers-color-scheme: dark) {
        .sp-copy-btn {
          color: #64b5f6;
        }
      }
    `;

    document.head.appendChild(style);
  }

  /**
   * Attach event listeners
   */
  private attachEventListeners(): void {
    // Close button
    document.addEventListener('click', (e) => {
      const target = e.target as HTMLElement;
      if (target.id === 'sp-inspector-close') {
        this.hide();
      } else if (target.id === 'sp-inspector-refresh') {
        this.refresh();
      } else if (
        target.classList.contains('sp-copy-btn') ||
        target.classList.contains('sp-copy-btn-inline') ||
        target.classList.contains('sp-value-copy-btn')
      ) {
        const value = target.getAttribute('data-value');
        if (value) {
          this.copyToClipboard(value);
        }
      } else if (target.classList.contains('sp-column-card')) {
        // Handle mouse click on column card to update selection
        const index = parseInt(target.getAttribute('data-column-index') || '-1', 10);
        if (index >= 0) {
          this.selectedColumnIndex = index;
          this.renderColumns();
        }
      }
    });

    // Search filter
    document.addEventListener('input', (e) => {
      const target = e.target as HTMLInputElement;
      if (target.id === 'sp-inspector-search') {
        this.filterText = target.value;
        this.renderColumns();
      }
    });

    // Show hidden checkbox
    document.addEventListener('change', (e) => {
      const target = e.target as HTMLInputElement;
      if (target.id === 'sp-inspector-show-hidden') {
        this.showHidden = target.checked;
        this.renderColumns();
      }
    });

    // Keyboard navigation
    document.addEventListener('keydown', (e) => {
      // Only handle keyboard events when inspector is visible
      if (!this.visible) return;

      const target = e.target as HTMLElement;
      const isSearchInput = target.id === 'sp-inspector-search' || target.tagName === 'INPUT';

      const columnsEl = document.getElementById('sp-inspector-columns');
      if (!columnsEl) return;

      const columnCards = Array.from(columnsEl.querySelectorAll('.sp-column-card'));
      const totalColumns = columnCards.length;

      if (totalColumns === 0) return;

      let handled = false;

      // Inner navigation mode - navigating within a column card
      if (this.innerNavigationMode) {
        handled = this.handleInnerNavigation(e, columnCards);
        if (handled) {
          e.preventDefault();
          e.stopPropagation();
        }
        return;
      }

      // Column-level navigation
      switch (e.key) {
        case 'ArrowDown':
          // Move to next column
          this.selectedColumnIndex = Math.min(this.selectedColumnIndex + 1, totalColumns - 1);
          handled = true;
          break;

        case 'ArrowUp':
          // Move to previous column
          if (this.selectedColumnIndex < 0) {
            this.selectedColumnIndex = 0;
          } else {
            this.selectedColumnIndex = Math.max(this.selectedColumnIndex - 1, 0);
          }
          handled = true;
          break;

        case 'Home':
          // Jump to first column
          this.selectedColumnIndex = 0;
          handled = true;
          break;

        case 'End':
          // Jump to last column
          this.selectedColumnIndex = totalColumns - 1;
          handled = true;
          break;

        case 'PageDown':
          this.selectedColumnIndex = Math.min(this.selectedColumnIndex + 10, totalColumns - 1);
          handled = true;
          break;

        case 'PageUp':
          this.selectedColumnIndex = Math.max(this.selectedColumnIndex - 10, 0);
          handled = true;
          break;

        case 'Enter':
          // Enter inner navigation mode to focus on buttons/links
          if (this.selectedColumnIndex >= 0 && this.selectedColumnIndex < totalColumns) {
            this.enterInnerNavigationMode(columnCards[this.selectedColumnIndex] as HTMLElement);
            handled = true;
          }
          break;

        case 'Escape':
          // Close inspector
          this.hide();
          handled = true;
          break;

        case '/':
          // Focus search input (only if not already in it)
          if (!isSearchInput) {
            const searchInput = document.getElementById('sp-inspector-search') as HTMLInputElement;
            if (searchInput) {
              searchInput.focus();
              searchInput.select();
              handled = true;
            }
          }
          break;
      }

      if (handled) {
        e.preventDefault();
        e.stopPropagation();
        this.renderColumns();
        this.scrollToSelectedColumn();
      }
    });
  }

  /**
   * Show the inspector and load data
   */
  async show(): Promise<void> {
    if (!this.container) return;

    this.container.style.display = 'flex';
    this.visible = true;

    // Reset selection when opening
    this.selectedColumnIndex = 0;

    await this.loadData();

    // Focus on the search input after a short delay to ensure DOM is ready
    setTimeout(() => {
      const searchInput = document.getElementById('sp-inspector-search') as HTMLInputElement;
      if (searchInput) {
        searchInput.focus();
        searchInput.select(); // Select existing text if any
      }
    }, 100);
  }

  /**
   * Hide the inspector
   */
  hide(): void {
    if (!this.container) return;

    this.container.style.display = 'none';
    this.visible = false;
  }

  /**
   * Toggle visibility
   */
  toggle(): void {
    if (this.visible) {
      this.hide();
    } else {
      this.show();
    }
  }

  /**
   * Refresh the data
   */
  async refresh(): Promise<void> {
    await this.loadData();
  }

  /**
   * Load list schema data
   */
  /**
   * Extract item ID or path from URL
   * Returns either a numeric ID or a server relative path
   */
  private extractItemIdOrPath(): number | string | null {
    const urlParams = new URLSearchParams(window.location.search);
    const idParam = urlParams.get('ID') || urlParams.get('id');

    console.log('[Column Inspector] Extracting item ID/path from URL:', window.location.href);
    console.log('[Column Inspector] ID parameter found:', idParam);

    if (idParam) {
      // Try parsing as numeric ID first
      const numericId = parseInt(idParam, 10);
      if (!isNaN(numericId)) {
        console.log('[Column Inspector] Valid numeric item ID extracted:', numericId);
        return numericId;
      }

      // Check if it's a file/folder path (starts with /)
      if (idParam.startsWith('/')) {
        console.log('[Column Inspector] Valid file/folder path extracted:', idParam);
        return idParam;
      }
    }

    console.log('[Column Inspector] No valid item ID or path found in URL');
    return null;
  }

  /**
   * Fetch item data from SharePoint REST API
   * Supports both numeric IDs and file/folder paths
   */
  private async fetchItemData(
    listId: string,
    itemIdOrPath: number | string,
    webUrl: string
  ): Promise<Record<string, any> | null> {
    try {
      let endpoint: string;

      if (typeof itemIdOrPath === 'number') {
        // Numeric ID - use standard items endpoint
        endpoint = `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemIdOrPath})`;
      } else {
        // File/folder path - try file first, then folder
        endpoint = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${itemIdOrPath}')/ListItemAllFields`;
      }

      console.log('[Column Inspector] Fetching item data from:', endpoint);

      const response = await fetch(endpoint, {
        method: 'GET',
        headers: {
          Accept: 'application/json;odata=verbose',
        },
        credentials: 'include',
      });

      console.log('[Column Inspector] Fetch response status:', response.status);

      if (!response.ok) {
        // If file endpoint failed with path, try folder endpoint
        if (typeof itemIdOrPath === 'string' && response.status === 404) {
          console.log('[Column Inspector] File endpoint failed, trying folder endpoint...');
          const folderEndpoint = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${itemIdOrPath}')/ListItemAllFields`;

          const folderResponse = await fetch(folderEndpoint, {
            method: 'GET',
            headers: {
              Accept: 'application/json;odata=verbose',
            },
            credentials: 'include',
          });

          console.log('[Column Inspector] Folder fetch response status:', folderResponse.status);

          if (!folderResponse.ok) {
            console.warn(
              '[Column Inspector] Failed to fetch folder data:',
              folderResponse.status,
              folderResponse.statusText
            );
            return null;
          }

          const folderData = await folderResponse.json();
          console.log('[Column Inspector] Folder item data received:', folderData.d);
          return folderData.d || null;
        }

        console.warn(
          '[Column Inspector] Failed to fetch item data:',
          response.status,
          response.statusText
        );
        return null;
      }

      const data = await response.json();
      console.log('[Column Inspector] Item data received:', data.d);
      return data.d || null;
    } catch (error) {
      console.error('[Column Inspector] Error fetching item data:', error);
      return null;
    }
  }

  private async loadData(): Promise<void> {
    const loadingEl = document.getElementById('sp-inspector-loading');
    const errorEl = document.getElementById('sp-inspector-error');
    const listInfoEl = document.getElementById('sp-inspector-list-info');
    const columnsEl = document.getElementById('sp-inspector-columns');

    if (!loadingEl || !errorEl || !listInfoEl || !columnsEl) return;

    // Show loading
    loadingEl.style.display = 'block';
    errorEl.style.display = 'none';
    listInfoEl.style.display = 'none';
    columnsEl.style.display = 'none';

    try {
      // Check if we're on a list view
      if (!ColumnFetcher.isListView()) {
        throw new Error('Not on a list/library view page');
      }

      // Get list ID using async extraction (supports classic SharePoint pages)
      loadingEl.textContent = 'Detecting list ID...';
      const listId = await extractListIdAsync();

      if (!listId) {
        throw new Error(
          'Could not detect list ID from current page. This may not be a SharePoint list/library page, or the list name could not be extracted from the URL.'
        );
      }

      // Fetch schema
      loadingEl.textContent = 'Loading column schema...';
      this.currentList = await ColumnFetcher.fetchListSchema(listId);

      // Try to detect and fetch current item data
      this.currentItemIdOrPath = this.extractItemIdOrPath();
      console.log('[Column Inspector] Current item ID/path:', this.currentItemIdOrPath);
      console.log('[Column Inspector] Web absolute URL:', this.currentList.webAbsoluteUrl);

      if (this.currentItemIdOrPath && this.currentList.webAbsoluteUrl) {
        const displayId =
          typeof this.currentItemIdOrPath === 'number'
            ? `#${this.currentItemIdOrPath}`
            : this.currentItemIdOrPath.split('/').pop();
        loadingEl.textContent = `Loading item ${displayId} data...`;
        this.currentItem = await this.fetchItemData(
          listId,
          this.currentItemIdOrPath,
          this.currentList.webAbsoluteUrl
        );
        console.log(
          '[Column Inspector] Current item data after fetch:',
          this.currentItem ? 'Data received' : 'No data'
        );
      } else {
        this.currentItem = null;
        console.log('[Column Inspector] Skipping item data fetch - no item ID/path or web URL');
      }

      // Hide loading, show data
      loadingEl.style.display = 'none';
      this.updateHeader();
      this.renderListInfo();
      this.renderColumns();
    } catch (error: unknown) {
      loadingEl.style.display = 'none';
      errorEl.style.display = 'block';

      // Provide helpful error messages
      let errorMessage = 'Error: ';
      if (error instanceof Error) {
        errorMessage += error.message;

        // Add helpful context for common errors
        if (error.message.includes('HTTP 404')) {
          errorMessage +=
            "\n\nThe list was not found. This may occur if:\n• The list has been deleted or renamed\n• You don't have access to the list\n• The URL format is not recognized";
        } else if (error.message.includes('HTTP 403')) {
          errorMessage +=
            "\n\nAccess denied. You may not have permission to view this list's schema.";
        } else if (error.message.includes('Could not detect list ID')) {
          errorMessage +=
            '\n\nTip: The column inspector works best on SharePoint list and library pages.';
        }
      } else {
        errorMessage += String(error);
      }

      errorEl.textContent = errorMessage;
      console.error('Column Inspector error:', error);
    }
  }

  /**
   * Render list information
   */
  private renderListInfo(): void {
    const listInfoEl = document.getElementById('sp-inspector-list-info');
    if (!listInfoEl || !this.currentList) return;

    const webUrl = this.currentList.webAbsoluteUrl || '';
    const listId = this.currentList.id;

    // Generate management links
    const listSettingsUrl = webUrl ? `${webUrl}/_layouts/15/listedit.aspx?List={${listId}}` : '';
    const contentTypesUrl = webUrl
      ? `${webUrl}/_layouts/15/ManageContentType.aspx?List={${listId}}`
      : '';

    // Build management links section
    let managementLinks = '';
    if (webUrl) {
      managementLinks = `
        <div class="sp-management-links">
          <a href="${sanitizeUrl(listSettingsUrl)}" target="_blank" rel="noopener noreferrer" class="sp-link-btn" title="Open list settings">
            ⚙️ List Settings
          </a>
          ${
            this.currentList.contentTypesEnabled
              ? `
            <a href="${sanitizeUrl(contentTypesUrl)}" target="_blank" rel="noopener noreferrer" class="sp-link-btn" title="Manage content types for this list">
              📋 Manage Content Types
            </a>
          `
              : ''
          }
        </div>
      `;
    }

    listInfoEl.style.display = 'block';
    listInfoEl.innerHTML = `
      <h4>${this.escapeHtml(this.currentList.title)}</h4>
      ${managementLinks}
      <div class="sp-info-row">
        <span class="sp-info-label">Internal Name:</span>
        <span class="sp-info-value">
          ${this.escapeHtml(this.currentList.internalName)}
          <button class="sp-copy-btn" data-value="${this.escapeHtml(this.currentList.internalName)}">Copy</button>
        </span>
      </div>
      <div class="sp-info-row">
        <span class="sp-info-label">List ID:</span>
        <span class="sp-info-value">
          ${this.escapeHtml(this.currentList.id)}
          <button class="sp-copy-btn" data-value="${this.escapeHtml(this.currentList.id)}">Copy</button>
        </span>
      </div>
      <div class="sp-info-row">
        <span class="sp-info-label">Template:</span>
        <span class="sp-info-value">${this.escapeHtml(this.currentList.baseType)} (${this.currentList.baseTemplate})</span>
      </div>
      <div class="sp-info-row">
        <span class="sp-info-label">URL:</span>
        <span class="sp-info-value">${this.escapeHtml(this.currentList.serverRelativeUrl)}</span>
      </div>
      <div class="sp-info-row">
        <span class="sp-info-label">Total Columns:</span>
        <span class="sp-info-value">${this.currentList.fields.length}</span>
      </div>
      ${
        this.currentList.contentTypesEnabled !== undefined
          ? `
        <div class="sp-info-row">
          <span class="sp-info-label">Content Types:</span>
          <span class="sp-info-value">${this.currentList.contentTypesEnabled ? '✓ Enabled' : '✗ Disabled'}</span>
        </div>
      `
          : ''
      }
      ${
        this.currentList.contentTypes && this.currentList.contentTypes.length > 0
          ? `
        <div class="sp-info-row sp-content-types-list">
          <span class="sp-info-label">Available Types:</span>
          <span class="sp-info-value sp-content-types-links">
            ${this.currentList.contentTypes
              .map(
                (ct) =>
                  `<a href="${sanitizeUrl(`${webUrl}/_layouts/15/ManageContentType.aspx?List={${listId}}&ctype=${encodeURIComponent(ct.id)}`)}" target="_blank" rel="noopener noreferrer" class="sp-content-type-link" title="${this.escapeHtml(ct.description || ct.name)}">${this.escapeHtml(ct.name)}</a>`
              )
              .join('')}
          </span>
        </div>
      `
          : ''
      }
    `;
  }

  /**
   * Update header with list name
   */
  private updateHeader(): void {
    const titleEl = document.getElementById('sp-inspector-title');
    const listNameEl = document.getElementById('sp-inspector-list-name');
    if (!titleEl || !listNameEl || !this.currentList) return;

    // Update the h3 to show the list title
    titleEl.textContent = this.currentList.title;

    // Show internal name in subtitle (only if different from title or always for clarity)
    const internalName = this.currentList.internalName;
    const showInternalName = internalName !== this.currentList.title || true; // Always show for PnP scripting

    if (showInternalName) {
      listNameEl.style.display = 'flex';
      listNameEl.innerHTML = `
        <span class="sp-list-internal-label">Internal Name:</span>
        <span class="sp-list-internal-value">${this.escapeHtml(internalName)}</span>
        <button class="sp-copy-btn-inline" data-value="${this.escapeHtml(internalName)}" title="Copy internal name for PnP scripting">Copy</button>
      `;
    }
  }

  /**
   * Render columns grid
   */
  private renderColumns(): void {
    const columnsEl = document.getElementById('sp-inspector-columns');
    if (!columnsEl || !this.currentList) return;

    // Filter columns
    let filteredColumns = this.currentList.fields;

    if (!this.showHidden) {
      filteredColumns = filteredColumns.filter((col) => !col.hidden);
    }

    if (this.filterText) {
      const searchLower = this.filterText.toLowerCase();
      filteredColumns = filteredColumns.filter(
        (col) =>
          col.title.toLowerCase().includes(searchLower) ||
          col.internalName.toLowerCase().includes(searchLower) ||
          col.type.toLowerCase().includes(searchLower)
      );
    }

    // Adjust selected index if it's out of bounds
    if (this.selectedColumnIndex >= filteredColumns.length) {
      this.selectedColumnIndex = filteredColumns.length - 1;
    }

    columnsEl.style.display = 'block';
    columnsEl.innerHTML = filteredColumns
      .map((col, index) => this.renderColumnCard(col, index))
      .join('');
  }

  /**
   * Render individual column card
   */
  private renderColumnCard(column: ColumnSchema, index: number): string {
    const isSelected = index === this.selectedColumnIndex;
    const selectedClass = isSelected ? 'selected' : '';
    const badges: string[] = [];

    if (column.required) badges.push('<span class="sp-badge required">Required</span>');
    if (column.hidden) badges.push('<span class="sp-badge hidden">Hidden</span>');
    if (column.readOnly) badges.push('<span class="sp-badge readonly">Read-only</span>');
    if (column.indexed) badges.push('<span class="sp-badge indexed">Indexed</span>');

    // Generate Edit Column link
    const webUrl = this.currentList?.webAbsoluteUrl || '';
    const listId = this.currentList?.id || '';
    const editColumnUrl =
      webUrl && listId
        ? `${webUrl}/_layouts/15/FldEdit.aspx?List={${listId}}&Field=${encodeURIComponent(column.internalName)}`
        : '';

    const details: string[] = [
      `<span class="sp-detail-label">Internal Name:</span>
       <span class="sp-detail-value">
         ${this.escapeHtml(column.internalName)}
         <button class="sp-copy-btn" data-value="${this.escapeHtml(column.internalName)}">Copy</button>
       </span>`,
      `<span class="sp-detail-label">Type:</span>
       <span class="sp-detail-value">${this.escapeHtml(column.type)}</span>`,
      `<span class="sp-detail-label">ID:</span>
       <span class="sp-detail-value">
         ${this.escapeHtml(column.id)}
         <button class="sp-copy-btn" data-value="${this.escapeHtml(column.id)}">Copy</button>
       </span>`,
    ];

    if (column.staticName) {
      details.push(
        `<span class="sp-detail-label">Static Name:</span>
         <span class="sp-detail-value">${this.escapeHtml(column.staticName)}</span>`
      );
    }

    if (column.group) {
      details.push(
        `<span class="sp-detail-label">Group:</span>
         <span class="sp-detail-value">${this.escapeHtml(column.group)}</span>`
      );
    }

    if (column.formula) {
      details.push(
        `<span class="sp-detail-label">Formula:</span>
         <span class="sp-detail-value">${this.escapeHtml(column.formula)}</span>`
      );
    }

    if (column.lookupListId) {
      details.push(
        `<span class="sp-detail-label">Lookup List:</span>
         <span class="sp-detail-value">${this.escapeHtml(column.lookupListId)}</span>`,
        `<span class="sp-detail-label">Lookup Field:</span>
         <span class="sp-detail-value">${this.escapeHtml(column.lookupField || '')}</span>`
      );
    }

    if (column.choices && column.choices.length > 0) {
      details.push(
        `<span class="sp-detail-label">Choices:</span>
         <span class="sp-detail-value">${this.escapeHtml(column.choices.join(', '))}</span>`
      );
    }

    if (column.maxLength) {
      details.push(
        `<span class="sp-detail-label">Max Length:</span>
         <span class="sp-detail-value">${column.maxLength}</span>`
      );
    }

    // Add current item value if available
    let currentValueSection = '';
    if (this.currentItem && this.currentItemIdOrPath) {
      const fieldValue = this.currentItem[column.internalName];
      const formattedValue = this.formatFieldValue(fieldValue, column.type);
      const displayId =
        typeof this.currentItemIdOrPath === 'number'
          ? `#${this.currentItemIdOrPath}`
          : this.currentItemIdOrPath.split('/').pop();

      currentValueSection = `
        <div class="sp-current-value-row">
          <div class="sp-current-value-label">Current Value (Item ${this.escapeHtml(displayId || '')})</div>
          <div class="sp-current-value-content">${formattedValue}</div>
          ${
            fieldValue !== null && fieldValue !== undefined
              ? `
          <div class="sp-current-value-actions">
            <button class="sp-value-copy-btn" data-value="${this.escapeHtml(String(fieldValue))}">📋 Copy Value</button>
          </div>
          `
              : ''
          }
        </div>
      `;
    }

    return `
      <div class="sp-column-card ${column.hidden ? 'hidden' : ''} ${selectedClass}" data-column-index="${index}" data-column-internal-name="${this.escapeHtml(column.internalName)}">
        <div class="sp-column-header">
          <div class="sp-column-title">${this.escapeHtml(column.title)}</div>
          <div class="sp-column-badges">${badges.join('')}</div>
        </div>
        ${
          this.currentItemIdOrPath
            ? `<div class="sp-item-indicator" style="background: #e1f5fe; padding: 4px 8px; font-size: 11px; color: #01579b; border-left: 3px solid #0078d4;">
                 📄 Viewing Item ${this.escapeHtml(typeof this.currentItemIdOrPath === 'number' ? `#${this.currentItemIdOrPath}` : this.currentItemIdOrPath.split('/').pop() || '')}
               </div>`
            : ''
        }
        ${
          editColumnUrl
            ? `
          <div class="sp-column-actions">
            <a href="${sanitizeUrl(editColumnUrl)}" target="_blank" rel="noopener noreferrer" class="sp-edit-column-link" title="Edit this column's settings">
              ✏️ Edit Column
            </a>
          </div>
        `
            : ''
        }
        <div class="sp-column-details">
          ${details.join('')}
          ${currentValueSection}
        </div>
      </div>
    `;
  }

  /**
   * Format field value for display based on type
   */
  private formatFieldValue(value: any, fieldType: string): string {
    // Handle null/undefined
    if (value === null || value === undefined) {
      return '<em style="color: #999;">(empty)</em>';
    }

    // Handle different field types
    switch (fieldType) {
      case 'Boolean':
        return value ? '✓ Yes' : '✗ No';

      case 'DateTime':
        try {
          const date = new Date(value);
          return this.escapeHtml(date.toLocaleString());
        } catch {
          return this.escapeHtml(String(value));
        }

      case 'User':
      case 'UserMulti':
        if (value.__deferred) {
          return '<em style="color: #999;">(User - expand to see)</em>';
        }
        if (typeof value === 'object' && value.Title) {
          return this.escapeHtml(value.Title);
        }
        return this.escapeHtml(String(value));

      case 'Lookup':
      case 'LookupMulti':
        if (value.__deferred) {
          return '<em style="color: #999;">(Lookup - expand to see)</em>';
        }
        if (typeof value === 'object') {
          const lookupValue = value.Title || value.LookupValue || JSON.stringify(value);
          return this.escapeHtml(String(lookupValue));
        }
        return this.escapeHtml(String(value));

      case 'URL':
        if (typeof value === 'object') {
          const url = value.Url || value.url || '';
          const desc = value.Description || value.description || url;
          const safeUrl = sanitizeUrl(url);
          if (safeUrl) {
            return `<a href="${safeUrl}" target="_blank" rel="noopener noreferrer" style="color: #0078d4;">${this.escapeHtml(desc)}</a>`;
          }
          // Unsafe scheme (javascript:, data:, ...) — show the raw text instead of a link
          if (url) {
            return this.escapeHtml(String(desc));
          }
        }
        return this.escapeHtml(String(value));

      case 'MultiChoice':
      case 'TaxonomyFieldTypeMulti':
        if (Array.isArray(value)) {
          return this.escapeHtml(value.join(', '));
        }
        if (typeof value === 'object' && value.results) {
          return this.escapeHtml(value.results.join(', '));
        }
        return this.escapeHtml(String(value));

      case 'Note': {
        // Multiline text - limit length
        const textValue = String(value);
        if (textValue.length > 200) {
          return this.escapeHtml(textValue.substring(0, 200)) + '... <em>(truncated)</em>';
        }
        return this.escapeHtml(textValue);
      }

      case 'Number':
      case 'Currency':
        if (typeof value === 'number') {
          return fieldType === 'Currency' ? `$${value.toFixed(2)}` : value.toString();
        }
        return this.escapeHtml(String(value));

      default:
        // For all other types, stringify if object, otherwise escape HTML
        if (typeof value === 'object') {
          return this.escapeHtml(JSON.stringify(value));
        }
        return this.escapeHtml(String(value));
    }
  }

  /**
   * Scroll selected column into view
   */
  private scrollToSelectedColumn(): void {
    if (this.selectedColumnIndex < 0) return;

    const columnsEl = document.getElementById('sp-inspector-columns');
    if (!columnsEl) return;

    const selectedCard = columnsEl.querySelector(
      `.sp-column-card[data-column-index="${this.selectedColumnIndex}"]`
    ) as HTMLElement;

    if (selectedCard) {
      selectedCard.scrollIntoView({
        behavior: 'smooth',
        block: 'nearest',
      });
    }
  }

  /**
   * Get focusable elements within a column card
   */
  private getFocusableElements(card: HTMLElement): HTMLElement[] {
    return Array.from(
      card.querySelectorAll('a.sp-edit-column-link, button.sp-copy-btn')
    ) as HTMLElement[];
  }

  /**
   * Enter inner navigation mode for a column card
   */
  private enterInnerNavigationMode(card: HTMLElement): void {
    const focusable = this.getFocusableElements(card);
    if (focusable.length === 0) return;

    this.innerNavigationMode = true;
    this.innerFocusIndex = 0;
    card.classList.add('inner-nav-active');
    // Force reflow before focus to ensure CSS is applied
    void card.offsetHeight;
    focusable[0].focus();
  }

  /**
   * Exit inner navigation mode
   */
  private exitInnerNavigationMode(): void {
    this.innerNavigationMode = false;
    this.innerFocusIndex = -1;

    // Remove inner-nav-active class from any card
    const activeCard = document.querySelector('.sp-column-card.inner-nav-active');
    if (activeCard) {
      activeCard.classList.remove('inner-nav-active');
    }

    // Re-render to restore focus indication
    this.renderColumns();
    this.scrollToSelectedColumn();
  }

  /**
   * Handle keyboard navigation within a column card
   */
  private handleInnerNavigation(e: KeyboardEvent, columnCards: Element[]): boolean {
    const selectedCard = columnCards[this.selectedColumnIndex] as HTMLElement;
    if (!selectedCard) return false;

    const focusable = this.getFocusableElements(selectedCard);
    if (focusable.length === 0) return false;

    switch (e.key) {
      case 'ArrowDown':
        // Move to next element
        this.innerFocusIndex = Math.min(this.innerFocusIndex + 1, focusable.length - 1);
        focusable[this.innerFocusIndex].focus();
        return true;

      case 'ArrowUp':
        // Move to previous element
        this.innerFocusIndex = Math.max(this.innerFocusIndex - 1, 0);
        focusable[this.innerFocusIndex].focus();
        return true;

      case 'Enter':
        // Activate the focused element
        focusable[this.innerFocusIndex]?.click();
        return true;

      case 'Escape':
        // Exit inner navigation mode
        this.exitInnerNavigationMode();
        return true;

      case 'Tab':
        // Prevent tab from leaving the card, cycle through elements
        if (e.shiftKey) {
          this.innerFocusIndex = Math.max(this.innerFocusIndex - 1, 0);
        } else {
          this.innerFocusIndex = Math.min(this.innerFocusIndex + 1, focusable.length - 1);
        }
        focusable[this.innerFocusIndex].focus();
        return true;
    }

    return false;
  }

  /**
   * Copy text to clipboard
   */
  private async copyToClipboard(text: string): Promise<void> {
    try {
      await navigator.clipboard.writeText(text);
      console.log('Copied to clipboard:', text);
    } catch (error) {
      console.error('Failed to copy:', error);
    }
  }

  /**
   * Escape HTML to prevent XSS
   */
  private escapeHtml(text: string): string {
    return escapeHtmlValue(text);
  }
}
