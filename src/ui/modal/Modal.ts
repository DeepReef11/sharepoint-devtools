/**
 * SharePoint DevTools Modal Component
 * A keyboard-driven modal for quick navigation
 */

import { ModalConfig, ResultItem, CategoryGroup, ModalState, ErrorDisplayOptions } from './types';
import { escapeHtml as escapeHtmlValue } from '../../utils/html';
import './modal.css';

export class Modal {
  private overlay: HTMLElement | null = null;
  private modal: HTMLElement | null = null;
  private searchInput: HTMLInputElement | null = null;
  private resultsContainer: HTMLElement | null = null;
  private stateContainer: HTMLElement | null = null;

  private config: ModalConfig;
  private isVisible: boolean = false;
  private results: ResultItem[] = [];
  private selectedIndex: number = -1;
  private flatResults: ResultItem[] = [];
  private currentState: ModalState = 'idle';
  private keyboardNavigationActive: boolean = false;

  private searchTimeout: number | null = null;
  private readonly SEARCH_DEBOUNCE_MS = 300;

  constructor(config: ModalConfig = {}) {
    this.config = {
      title: 'SharePoint DevTools',
      placeholder: 'Search for SharePoint locations...',
      darkMode: false,
      ...config,
    };

    this.injectStyles();
    this.createModal();
    this.attachEventListeners();
  }

  /**
   * Inject modal styles into the page
   * Styles are now imported via webpack, so this is a no-op
   */
  private injectStyles(): void {
    // Styles are imported at the top of this file via webpack
    // No need to inject them manually
  }

  /**
   * Create modal DOM structure
   */
  private createModal(): void {
    // Create overlay
    this.overlay = document.createElement('div');
    this.overlay.className = 'spqn-overlay';
    if (this.config.darkMode) {
      this.overlay.classList.add('spqn-dark-mode');
    }

    // Create modal container
    this.modal = document.createElement('div');
    this.modal.className = 'spqn-modal';

    // Create header
    const header = document.createElement('div');
    header.className = 'spqn-header';
    const title = document.createElement('h1');
    title.className = 'spqn-title';
    title.textContent = this.config.title || '';
    header.appendChild(title);

    // Create search container
    const searchContainer = document.createElement('div');
    searchContainer.className = 'spqn-search-container';
    this.searchInput = document.createElement('input');
    this.searchInput.type = 'text';
    this.searchInput.className = 'spqn-search-input';
    this.searchInput.placeholder = this.config.placeholder || '';
    this.searchInput.autocomplete = 'off';
    this.searchInput.spellcheck = false;
    searchContainer.appendChild(this.searchInput);

    // Create results container
    this.resultsContainer = document.createElement('div');
    this.resultsContainer.className = 'spqn-results-container';

    // Create state container (for loading, error, empty states)
    this.stateContainer = document.createElement('div');
    this.stateContainer.className = 'spqn-state-container spqn-hidden';

    // Create footer
    const footer = document.createElement('div');
    footer.className = 'spqn-footer';
    footer.innerHTML = `
      <div class="spqn-keyboard-hint">
        <span><span class="spqn-key">↑</span><span class="spqn-key">↓</span> Navigate</span>
        <span><span class="spqn-key">PgUp</span><span class="spqn-key">PgDn</span> Page</span>
        <span><span class="spqn-key">Home</span><span class="spqn-key">End</span> Jump</span>
        <span><span class="spqn-key">Enter</span> Select</span>
        <span><span class="spqn-key">Esc</span> Close</span>
      </div>
      <div>v1.0.0</div>
    `;

    // Assemble modal
    this.modal.appendChild(header);
    this.modal.appendChild(searchContainer);
    this.modal.appendChild(this.resultsContainer);
    this.modal.appendChild(this.stateContainer);
    this.modal.appendChild(footer);

    this.overlay.appendChild(this.modal);
  }

  /**
   * Attach event listeners
   */
  private attachEventListeners(): void {
    // Close on overlay click
    this.overlay?.addEventListener('click', (e) => {
      if (e.target === this.overlay) {
        this.close();
      }
    });

    // Prevent modal click from closing
    this.modal?.addEventListener('click', (e) => {
      e.stopPropagation();
    });

    // Search input events
    this.searchInput?.addEventListener('input', (e) => {
      const query = (e.target as HTMLInputElement).value;
      this.handleSearch(query);
    });

    // Keyboard navigation
    this.searchInput?.addEventListener('keydown', (e) => {
      this.handleKeyDown(e);
    });

    // Keyboard-only selection mode (mouse hover disabled)
    this.keyboardNavigationActive = true;

    // Global keyboard shortcut (Escape)
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && this.isVisible) {
        this.close();
      }
    });
  }

  /**
   * Handle search input with debouncing
   */
  private handleSearch(query: string): void {
    if (this.searchTimeout) {
      window.clearTimeout(this.searchTimeout);
    }

    this.searchTimeout = window.setTimeout(() => {
      if (this.config.onSearch) {
        this.config.onSearch(query);
      }
    }, this.SEARCH_DEBOUNCE_MS);
  }

  /**
   * Handle keyboard navigation
   */
  private handleKeyDown(e: KeyboardEvent): void {
    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        this.keyboardNavigationActive = true;
        this.selectNext();
        break;

      case 'ArrowUp':
        e.preventDefault();
        this.keyboardNavigationActive = true;
        this.selectPrevious();
        break;

      case 'PageDown':
        e.preventDefault();
        this.keyboardNavigationActive = true;
        this.pageDown();
        break;

      case 'PageUp':
        e.preventDefault();
        this.keyboardNavigationActive = true;
        this.pageUp();
        break;

      case 'Enter':
        e.preventDefault();
        this.selectCurrent();
        break;

      case 'Home':
        e.preventDefault();
        this.keyboardNavigationActive = true;
        this.selectFirst();
        break;

      case 'End':
        e.preventDefault();
        this.keyboardNavigationActive = true;
        this.selectLast();
        break;
    }
  }

  /**
   * Navigate to next item
   */
  private selectNext(): void {
    if (this.flatResults.length === 0) return;

    this.selectedIndex = Math.min(this.selectedIndex + 1, this.flatResults.length - 1);
    this.updateSelection();
  }

  /**
   * Navigate to previous item
   */
  private selectPrevious(): void {
    if (this.flatResults.length === 0) return;

    this.selectedIndex = Math.max(this.selectedIndex - 1, 0);
    this.updateSelection();
  }

  /**
   * Select first item
   */
  private selectFirst(): void {
    if (this.flatResults.length === 0) return;

    this.selectedIndex = 0;
    this.updateSelection();
  }

  /**
   * Select last item
   */
  private selectLast(): void {
    if (this.flatResults.length === 0) return;

    this.selectedIndex = this.flatResults.length - 1;
    this.updateSelection();
  }

  /**
   * Navigate down by one page
   */
  private pageDown(): void {
    if (this.flatResults.length === 0) return;

    const pageSize = this.calculatePageSize();
    this.selectedIndex = Math.min(this.selectedIndex + pageSize, this.flatResults.length - 1);
    this.updateSelection();
  }

  /**
   * Navigate up by one page
   */
  private pageUp(): void {
    if (this.flatResults.length === 0) return;

    const pageSize = this.calculatePageSize();
    this.selectedIndex = Math.max(this.selectedIndex - pageSize, 0);
    this.updateSelection();
  }

  /**
   * Calculate the number of items visible in one page
   * Based on container height and item height
   */
  private calculatePageSize(): number {
    if (!this.resultsContainer) return 10; // Default fallback

    const containerHeight = this.resultsContainer.clientHeight;
    const firstItem = this.resultsContainer.querySelector('.spqn-result-item');

    if (!firstItem) return 10; // Default fallback

    const itemHeight = firstItem.getBoundingClientRect().height;

    // Calculate visible items, minimum of 1
    return Math.max(1, Math.floor(containerHeight / itemHeight));
  }

  /**
   * Update visual selection
   * Optimized to minimize layout thrashing
   */
  private updateSelection(): void {
    const items = this.resultsContainer?.querySelectorAll('.spqn-result-item');
    if (!items) return;

    // Batch read operations first (avoid layout thrashing)
    let selectedElement: Element | null = null;

    // Then batch write operations
    items.forEach((item, index) => {
      if (index === this.selectedIndex) {
        item.classList.add('spqn-selected');
        selectedElement = item;
      } else {
        item.classList.remove('spqn-selected');
      }
    });

    // Scroll selected element into view (causes layout)
    if (selectedElement) {
      selectedElement.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
  }

  /**
   * Select current item and trigger callback
   */
  private selectCurrent(): void {
    if (this.selectedIndex >= 0 && this.selectedIndex < this.flatResults.length) {
      const selectedItem = this.flatResults[this.selectedIndex];

      if (this.config.onSelect) {
        this.config.onSelect(selectedItem);
      }

      // If item has a URL, navigate to it
      if (selectedItem.url) {
        window.location.href = selectedItem.url;
      }

      this.close();
    }
  }

  /**
   * Show modal
   */
  public show(): void {
    if (this.isVisible || !this.overlay) return;

    document.body.appendChild(this.overlay);
    this.isVisible = true;

    // Focus search input and trigger initial search
    window.setTimeout(() => {
      this.searchInput?.focus();
      // Trigger initial search with empty query to show default results
      if (this.searchInput) {
        this.searchInput.value = '';
        this.handleSearch('');
      }
    }, 100);
  }

  /**
   * Close modal
   */
  public close(): void {
    if (!this.isVisible || !this.overlay) return;

    // Add closing animation
    this.overlay.classList.add('spqn-closing');

    window.setTimeout(() => {
      if (this.overlay && this.overlay.parentNode) {
        this.overlay.parentNode.removeChild(this.overlay);
      }
      this.overlay?.classList.remove('spqn-closing');
      this.isVisible = false;

      // Reset state
      if (this.searchInput) {
        this.searchInput.value = '';
      }
      this.selectedIndex = -1;
      this.results = [];
      this.flatResults = [];
      this.keyboardNavigationActive = false;

      if (this.config.onClose) {
        this.config.onClose();
      }
    }, 150);
  }

  /**
   * Toggle modal visibility
   */
  public toggle(): void {
    if (this.isVisible) {
      this.close();
    } else {
      this.show();
    }
  }

  /**
   * Set results and render
   */
  public setResults(results: ResultItem[]): void {
    this.results = results;
    this.flatResults = results;
    this.selectedIndex = results.length > 0 ? 0 : -1;
    this.renderResults();
  }

  /**
   * Set results with category grouping
   */
  public setGroupedResults(groups: CategoryGroup[]): void {
    this.results = [];
    this.flatResults = [];

    groups.forEach((group) => {
      this.results.push(...group.items);
      this.flatResults.push(...group.items);
    });

    this.selectedIndex = this.flatResults.length > 0 ? 0 : -1;
    this.renderGroupedResults(groups);
  }

  /**
   * Render results (flat list)
   * Optimized with DocumentFragment to minimize DOM reflows
   */
  private renderResults(): void {
    if (!this.resultsContainer) return;

    this.hideState();
    this.resultsContainer.innerHTML = '';

    if (this.results.length === 0) {
      this.showState('empty', 'No results found');
      return;
    }

    // Use DocumentFragment for batch DOM insertion
    const fragment = document.createDocumentFragment();
    this.results.forEach((item, index) => {
      const itemElement = this.createResultItem(item, index);
      fragment.appendChild(itemElement);
    });

    // Single DOM operation instead of multiple appendChild calls
    this.resultsContainer.appendChild(fragment);

    this.updateSelection();
  }

  /**
   * Render results with category grouping
   * Optimized with DocumentFragment to minimize DOM reflows
   */
  private renderGroupedResults(groups: CategoryGroup[]): void {
    if (!this.resultsContainer) return;

    this.hideState();
    this.resultsContainer.innerHTML = '';

    if (groups.length === 0 || this.flatResults.length === 0) {
      this.showState('empty', 'No results found');
      return;
    }

    // Use DocumentFragment for batch DOM insertion
    const fragment = document.createDocumentFragment();
    let globalIndex = 0;

    groups.forEach((group) => {
      if (group.items.length === 0) return;

      // Create category header
      const categoryHeader = document.createElement('div');
      categoryHeader.className = 'spqn-category-header';
      categoryHeader.textContent = group.category;
      fragment.appendChild(categoryHeader);

      // Create category group container
      const categoryGroup = document.createElement('div');
      categoryGroup.className = 'spqn-category-group';

      group.items.forEach((item) => {
        const itemElement = this.createResultItem(item, globalIndex);
        categoryGroup.appendChild(itemElement);
        globalIndex++;
      });

      fragment.appendChild(categoryGroup);
    });

    // Single DOM operation instead of multiple appendChild calls
    this.resultsContainer.appendChild(fragment);

    this.updateSelection();
  }

  /**
   * Create a result item element
   * Optimized to minimize DOM manipulation
   * Note: Mouse interaction is disabled - keyboard-only navigation
   */
  private createResultItem(item: ResultItem, index: number): HTMLElement {
    const itemElement = document.createElement('div');
    itemElement.className = 'spqn-result-item';
    itemElement.dataset.index = index.toString();

    // Build HTML structure in one go using template
    const titleDiv = `<div class="spqn-result-title">${this.escapeHtml(item.title)}</div>`;
    const descDiv = item.description
      ? `<div class="spqn-result-description">${this.escapeHtml(item.description)}</div>`
      : '';

    itemElement.innerHTML = titleDiv + descDiv;

    // Mouse interaction disabled - keyboard-only navigation mode
    // No click or hover handlers attached to enforce keyboard-only selection

    return itemElement;
  }

  /**
   * Escape HTML to prevent XSS
   */
  private escapeHtml(text: string): string {
    return escapeHtmlValue(text);
  }

  /**
   * Set modal state (loading, error, etc.)
   */
  public setState(state: ModalState, message?: string): void {
    this.currentState = state;

    switch (state) {
      case 'loading':
        this.showState('loading', message || 'Loading...');
        break;

      case 'error':
        this.showState('error', message || 'An error occurred');
        break;

      case 'empty':
        this.showState('empty', message || 'No results found');
        break;

      case 'idle':
      case 'success':
        this.hideState();
        break;
    }
  }

  /**
   * Show detailed error with optional retry
   */
  public showError(options: ErrorDisplayOptions): void {
    if (!this.stateContainer || !this.resultsContainer) return;

    this.resultsContainer.classList.add('spqn-hidden');
    this.stateContainer.classList.remove('spqn-hidden');
    this.stateContainer.className = 'spqn-state-container spqn-error';

    const title = options.title || 'Error';
    const retryButton = options.canRetry
      ? '<button class="spqn-retry-button">Try Again</button>'
      : '';

    this.stateContainer.innerHTML = `
      <div class="spqn-state-icon">⚠</div>
      <div class="spqn-error-title">${this.escapeHtml(title)}</div>
      <div class="spqn-state-message">${this.escapeHtml(options.message)}</div>
      ${options.details ? `<div class="spqn-error-details">${this.escapeHtml(options.details)}</div>` : ''}
      <div class="spqn-error-buttons">
        ${retryButton}
        <button class="spqn-close-button">Close</button>
      </div>
    `;

    if (options.canRetry && options.onRetry) {
      const button = this.stateContainer.querySelector('.spqn-retry-button');
      button?.addEventListener('click', () => {
        options.onRetry!();
      });
    }

    // Add close button handler
    const closeButton = this.stateContainer.querySelector('.spqn-close-button');
    closeButton?.addEventListener('click', () => {
      this.close();
    });
  }

  /**
   * Show state container
   */
  private showState(type: 'loading' | 'error' | 'empty', message: string): void {
    if (!this.stateContainer || !this.resultsContainer) return;

    this.resultsContainer.classList.add('spqn-hidden');
    this.stateContainer.classList.remove('spqn-hidden');
    this.stateContainer.className = `spqn-state-container spqn-${type}`;

    const icon = type === 'loading' ? '⟳' : type === 'error' ? '⚠' : '∅';

    this.stateContainer.innerHTML = `
      <div class="spqn-state-icon">${icon}</div>
      <div class="spqn-state-message">${this.escapeHtml(message)}</div>
    `;
  }

  /**
   * Hide state container
   */
  private hideState(): void {
    if (!this.stateContainer || !this.resultsContainer) return;

    this.stateContainer.classList.add('spqn-hidden');
    this.resultsContainer.classList.remove('spqn-hidden');
  }

  /**
   * Check if modal is visible
   */
  public isOpen(): boolean {
    return this.isVisible;
  }

  /**
   * Set dark mode
   */
  public setDarkMode(enabled: boolean): void {
    this.config.darkMode = enabled;

    if (this.overlay) {
      if (enabled) {
        this.overlay.classList.add('spqn-dark-mode');
      } else {
        this.overlay.classList.remove('spqn-dark-mode');
      }
    }
  }

  /**
   * Destroy modal and clean up
   */
  public destroy(): void {
    if (this.isVisible) {
      this.close();
    }

    if (this.overlay && this.overlay.parentNode) {
      this.overlay.parentNode.removeChild(this.overlay);
    }

    this.overlay = null;
    this.modal = null;
    this.searchInput = null;
    this.resultsContainer = null;
    this.stateContainer = null;
  }
}
