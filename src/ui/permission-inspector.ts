/**
 * Permission Inspector UI Component
 * Displays site permissions, groups, users, and role assignments in a side panel
 */

import {
  SPUser,
  SPGroup,
  PermissionLevel,
  PermissionInfo,
  getWebPermissions,
  getSiteGroups,
  getGroupUsers,
  getPermissionLevels,
  getCurrentUser,
  getSitePermissionsUrl,
  getGroupManagementUrl,
  getRoleTypeName,
} from '../api/permissions-api';
import { escapeHtml as escapeHtmlValue, sanitizeUrl } from '../utils/html';
import { getSharePointContext } from '../context/sharepoint-context';

type ViewMode = 'assignments' | 'groups' | 'levels';

export class PermissionInspector {
  private container: HTMLElement | null = null;
  private visible: boolean = false;
  private filterText: string = '';
  private selectedIndex: number = -1;
  private innerNavigationMode: boolean = false;
  private innerFocusIndex: number = -1;

  // Data
  private permissionInfo: PermissionInfo | null = null;
  private siteGroups: SPGroup[] = [];
  private permissionLevels: PermissionLevel[] = [];
  private currentUser: SPUser | null = null;
  private expandedGroups: Set<number> = new Set();
  private groupUsers: Map<number, SPUser[]> = new Map();

  // View state
  private viewMode: ViewMode = 'assignments';

  constructor() {
    this.createUI();
    this.attachEventListeners();
  }

  /**
   * Create the inspector UI
   */
  private createUI(): void {
    const inspector = document.createElement('div');
    inspector.id = 'sp-permission-inspector';
    inspector.className = 'sp-permission-inspector';
    inspector.style.display = 'none';

    inspector.innerHTML = `
      <div class="sp-inspector-header">
        <div class="sp-inspector-title-section">
          <h3 id="sp-perm-inspector-title">Permission Inspector</h3>
          <div id="sp-perm-inspector-subtitle" class="sp-inspector-list-name" style="display: none;"></div>
        </div>
        <div class="sp-inspector-controls">
          <button id="sp-perm-inspector-refresh" class="sp-btn" title="Refresh">
            🔄
          </button>
          <button id="sp-perm-inspector-close" class="sp-btn" title="Close">
            ✕
          </button>
        </div>
      </div>
      <div class="sp-inspector-toolbar">
        <input
          type="text"
          id="sp-perm-inspector-search"
          class="sp-search-input"
          placeholder="Filter permissions..."
        />
      </div>
      <div class="sp-perm-view-tabs">
        <button class="sp-perm-tab active" data-view="assignments">Role Assignments</button>
        <button class="sp-perm-tab" data-view="groups">Groups</button>
        <button class="sp-perm-tab" data-view="levels">Permission Levels</button>
      </div>
      <div class="sp-inspector-content">
        <div id="sp-perm-inspector-loading" class="sp-loading">
          Loading permissions...
        </div>
        <div id="sp-perm-inspector-error" class="sp-error" style="display: none;">
        </div>
        <div id="sp-perm-inspector-data" class="sp-perm-data" style="display: none;">
        </div>
      </div>
    `;

    this.container = inspector;
    document.body.appendChild(inspector);

    this.injectStyles();
  }

  /**
   * Inject CSS styles
   */
  private injectStyles(): void {
    if (document.getElementById('sp-perm-inspector-styles')) {
      return;
    }

    const style = document.createElement('style');
    style.id = 'sp-perm-inspector-styles';
    style.textContent = `
      .sp-permission-inspector {
        position: fixed;
        top: 0;
        right: 0;
        width: 520px;
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
        .sp-permission-inspector {
          background: #1e1e1e;
          color: #d4d4d4;
          box-shadow: -2px 0 10px rgba(0, 0, 0, 0.5);
        }
      }

      .sp-perm-view-tabs {
        display: flex;
        padding: 0 16px;
        border-bottom: 1px solid #e0e0e0;
        gap: 4px;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-view-tabs {
          border-bottom-color: #333;
        }
      }

      .sp-perm-tab {
        padding: 10px 16px;
        background: transparent;
        border: none;
        border-bottom: 2px solid transparent;
        cursor: pointer;
        font-size: 13px;
        font-weight: 500;
        color: #666;
        transition: all 0.2s;
      }

      .sp-perm-tab:hover {
        color: #333;
        background: #f5f5f5;
      }

      .sp-perm-tab.active {
        color: #0078d4;
        border-bottom-color: #0078d4;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-tab {
          color: #999;
        }
        .sp-perm-tab:hover {
          color: #d4d4d4;
          background: #2a2a2a;
        }
        .sp-perm-tab.active {
          color: #64b5f6;
          border-bottom-color: #64b5f6;
        }
      }

      .sp-perm-data {
        display: flex;
        flex-direction: column;
        gap: 12px;
      }

      .sp-perm-card {
        border: 1px solid #e0e0e0;
        border-radius: 6px;
        padding: 12px;
        background: #fafafa;
        cursor: pointer;
        transition: background 0.2s, border-color 0.2s, box-shadow 0.2s;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-card {
          background: #2a2a2a;
          border-color: #444;
        }
      }

      .sp-perm-card:hover {
        background: #f0f0f0;
        border-color: #ccc;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-card:hover {
          background: #333;
          border-color: #555;
        }
      }

      .sp-perm-card.selected {
        background: #e3f2fd;
        border-color: #1976d2;
        border-width: 2px;
        box-shadow: 0 0 0 2px rgba(25, 118, 210, 0.2);
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-card.selected {
          background: #1a3a52;
          border-color: #64b5f6;
          box-shadow: 0 0 0 2px rgba(100, 181, 246, 0.2);
        }
      }

      .sp-perm-card.inner-nav-active {
        background: #bbdefb;
        border-color: #0d47a1;
        border-width: 2px;
        box-shadow: 0 0 0 3px rgba(13, 71, 161, 0.3);
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-card.inner-nav-active {
          background: #1a4a6a;
          border-color: #90caf9;
          box-shadow: 0 0 0 3px rgba(144, 202, 249, 0.3);
        }
      }

      .sp-perm-card-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 8px;
      }

      .sp-perm-card-title {
        font-weight: 600;
        font-size: 15px;
        display: flex;
        align-items: center;
        gap: 8px;
      }

      .sp-perm-card-subtitle {
        font-size: 12px;
        color: #666;
        margin-top: 2px;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-card-subtitle {
          color: #999;
        }
      }

      .sp-perm-type-badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 500;
        text-transform: uppercase;
      }

      .sp-perm-type-badge.user {
        background: #e8f5e9;
        color: #2e7d32;
      }

      .sp-perm-type-badge.group {
        background: #e3f2fd;
        color: #1565c0;
      }

      .sp-perm-type-badge.admin {
        background: #fff3e0;
        color: #ef6c00;
      }

      .sp-perm-type-badge.hidden {
        background: #fafafa;
        color: #666;
        border: 1px solid #ddd;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-type-badge.user {
          background: #1b5e20;
          color: #a5d6a7;
        }
        .sp-perm-type-badge.group {
          background: #0d47a1;
          color: #90caf9;
        }
        .sp-perm-type-badge.admin {
          background: #e65100;
          color: #ffcc80;
        }
        .sp-perm-type-badge.hidden {
          background: #2a2a2a;
          color: #999;
          border-color: #444;
        }
      }

      .sp-perm-roles {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin-top: 8px;
      }

      .sp-perm-role-chip {
        display: inline-flex;
        align-items: center;
        padding: 4px 10px;
        background: #f0f0f0;
        border-radius: 4px;
        font-size: 12px;
        color: #333;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-role-chip {
          background: #3a3a3a;
          color: #d4d4d4;
        }
      }

      .sp-perm-details {
        display: grid;
        grid-template-columns: auto 1fr;
        gap: 6px 12px;
        font-size: 13px;
        margin-top: 10px;
      }

      .sp-perm-detail-label {
        color: #666;
        font-weight: 500;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-detail-label {
          color: #999;
        }
      }

      .sp-perm-detail-value {
        font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
        word-break: break-all;
      }

      .sp-perm-group-users {
        margin-top: 10px;
        padding-top: 10px;
        border-top: 1px solid #e0e0e0;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-group-users {
          border-top-color: #444;
        }
      }

      .sp-perm-group-users-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 8px;
      }

      .sp-perm-group-users-title {
        font-size: 12px;
        font-weight: 600;
        color: #666;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-group-users-title {
          color: #999;
        }
      }

      .sp-perm-expand-btn {
        background: none;
        border: 1px solid #ccc;
        border-radius: 4px;
        padding: 4px 10px;
        cursor: pointer;
        font-size: 11px;
        color: #666;
        transition: all 0.2s;
      }

      .sp-perm-expand-btn:hover {
        background: #f0f0f0;
        border-color: #999;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-expand-btn {
          border-color: #555;
          color: #999;
        }
        .sp-perm-expand-btn:hover {
          background: #333;
          border-color: #666;
        }
      }

      .sp-perm-user-list {
        display: flex;
        flex-direction: column;
        gap: 6px;
      }

      .sp-perm-user-item {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 6px 8px;
        background: #f5f5f5;
        border-radius: 4px;
        font-size: 12px;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-user-item {
          background: #333;
        }
      }

      .sp-perm-user-icon {
        font-size: 14px;
      }

      .sp-perm-user-name {
        flex: 1;
      }

      .sp-perm-user-email {
        color: #666;
        font-size: 11px;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-user-email {
          color: #999;
        }
      }

      .sp-perm-info-banner {
        padding: 12px;
        background: #e3f2fd;
        border-radius: 6px;
        margin-bottom: 12px;
        font-size: 13px;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-info-banner {
          background: #1a3a52;
        }
      }

      .sp-perm-info-banner a {
        color: #1976d2;
        text-decoration: none;
      }

      .sp-perm-info-banner a:hover {
        text-decoration: underline;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-info-banner a {
          color: #64b5f6;
        }
      }

      .sp-perm-link-btn {
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
        margin-top: 8px;
      }

      .sp-perm-link-btn:hover {
        background: #edebe9;
        border-color: #d2d0ce;
      }

      @media (prefers-color-scheme: dark) {
        .sp-perm-link-btn {
          background: #3a3a3a;
          color: #d4d4d4;
          border-color: #555;
        }
        .sp-perm-link-btn:hover {
          background: #4a4a4a;
          border-color: #666;
        }
      }
    `;

    document.head.appendChild(style);
  }

  /**
   * Attach event listeners
   */
  private attachEventListeners(): void {
    document.addEventListener('click', (e) => {
      const target = e.target as HTMLElement;

      if (target.id === 'sp-perm-inspector-close') {
        this.hide();
      } else if (target.id === 'sp-perm-inspector-refresh') {
        this.refresh();
      } else if (target.classList.contains('sp-perm-tab')) {
        const view = target.getAttribute('data-view') as ViewMode;
        if (view) {
          this.setViewMode(view);
        }
      } else if (target.classList.contains('sp-perm-expand-btn')) {
        const groupId = parseInt(target.getAttribute('data-group-id') || '0', 10);
        if (groupId) {
          this.toggleGroupExpansion(groupId);
        }
      } else if (target.classList.contains('sp-perm-card')) {
        const index = parseInt(target.getAttribute('data-index') || '-1', 10);
        if (index >= 0) {
          this.selectedIndex = index;
          this.renderData();
        }
      }
    });

    // Search filter
    document.addEventListener('input', (e) => {
      const target = e.target as HTMLInputElement;
      if (target.id === 'sp-perm-inspector-search') {
        this.filterText = target.value;
        this.renderData();
      }
    });

    // Keyboard navigation
    document.addEventListener('keydown', (e) => {
      if (!this.visible) return;

      const target = e.target as HTMLElement;
      const isSearchInput = target.id === 'sp-perm-inspector-search' || target.tagName === 'INPUT';

      const dataEl = document.getElementById('sp-perm-inspector-data');
      if (!dataEl) return;

      const cards = Array.from(dataEl.querySelectorAll('.sp-perm-card'));
      const totalCards = cards.length;

      if (totalCards === 0) return;

      let handled = false;

      if (this.innerNavigationMode) {
        handled = this.handleInnerNavigation(e, cards);
        if (handled) {
          e.preventDefault();
          e.stopPropagation();
        }
        return;
      }

      const tabs: ViewMode[] = ['assignments', 'groups', 'levels'];
      const currentTabIndex = tabs.indexOf(this.viewMode);

      switch (e.key) {
        case 'ArrowDown':
          this.selectedIndex = Math.min(this.selectedIndex + 1, totalCards - 1);
          handled = true;
          break;

        case 'ArrowUp':
          if (this.selectedIndex < 0) {
            this.selectedIndex = 0;
          } else {
            this.selectedIndex = Math.max(this.selectedIndex - 1, 0);
          }
          handled = true;
          break;

        case 'ArrowLeft':
          if (currentTabIndex > 0) {
            this.setViewMode(tabs[currentTabIndex - 1]);
            handled = true;
          }
          break;

        case 'ArrowRight':
          if (currentTabIndex < tabs.length - 1) {
            this.setViewMode(tabs[currentTabIndex + 1]);
            handled = true;
          }
          break;

        case 'Home':
          this.selectedIndex = 0;
          handled = true;
          break;

        case 'End':
          this.selectedIndex = totalCards - 1;
          handled = true;
          break;

        case 'PageDown':
          this.selectedIndex = Math.min(this.selectedIndex + 10, totalCards - 1);
          handled = true;
          break;

        case 'PageUp':
          this.selectedIndex = Math.max(this.selectedIndex - 10, 0);
          handled = true;
          break;

        case 'Enter':
          if (this.selectedIndex >= 0 && this.selectedIndex < totalCards) {
            // In groups view, directly toggle expansion instead of entering inner nav
            if (this.viewMode === 'groups') {
              const card = cards[this.selectedIndex] as HTMLElement;
              const groupId = card.dataset.groupId;
              if (groupId) {
                this.toggleGroupExpansion(parseInt(groupId, 10));
                handled = true;
              }
            } else {
              this.enterInnerNavigationMode(cards[this.selectedIndex] as HTMLElement);
              handled = true;
            }
          }
          break;

        case 'Escape':
          this.hide();
          handled = true;
          break;

        case '/':
          // Focus search input (only if not already in it)
          if (!isSearchInput) {
            const searchInput = document.getElementById(
              'sp-perm-inspector-search'
            ) as HTMLInputElement;
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
        this.renderData();
        this.scrollToSelectedCard();
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
    this.selectedIndex = 0;

    await this.loadData();

    setTimeout(() => {
      const searchInput = document.getElementById('sp-perm-inspector-search') as HTMLInputElement;
      if (searchInput) {
        searchInput.focus();
        searchInput.select();
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
   * Set view mode
   */
  private setViewMode(mode: ViewMode): void {
    this.viewMode = mode;
    this.selectedIndex = 0;

    // Update tab styling
    const tabs = document.querySelectorAll('.sp-perm-tab');
    tabs.forEach((tab) => {
      tab.classList.toggle('active', tab.getAttribute('data-view') === mode);
    });

    this.renderData();
  }

  /**
   * Load permission data
   */
  private async loadData(): Promise<void> {
    const loadingEl = document.getElementById('sp-perm-inspector-loading');
    const errorEl = document.getElementById('sp-perm-inspector-error');
    const dataEl = document.getElementById('sp-perm-inspector-data');

    if (!loadingEl || !errorEl || !dataEl) return;

    loadingEl.style.display = 'block';
    errorEl.style.display = 'none';
    dataEl.style.display = 'none';

    try {
      const context = getSharePointContext();

      // Get webUrl with fallback logic (same fix as content-script)
      let webUrl = context?.webUrl || '';
      const siteUrl = context?.siteUrl || '';

      // If siteUrl has /sites/ but webUrl doesn't, use siteUrl
      if (siteUrl.includes('/sites/') && !webUrl.includes('/sites/')) {
        webUrl = siteUrl;
      }

      // Final fallback to origin
      if (!webUrl) {
        webUrl = window.location.origin;
      }

      if (!webUrl) {
        throw new Error('Could not detect SharePoint context');
      }

      loadingEl.textContent = 'Loading permissions...';

      // Load all data in parallel
      const [permInfo, groups, levels, user] = await Promise.all([
        getWebPermissions(webUrl),
        getSiteGroups(webUrl),
        getPermissionLevels(webUrl),
        getCurrentUser(webUrl),
      ]);

      this.permissionInfo = permInfo;
      this.siteGroups = groups;
      this.permissionLevels = levels;
      this.currentUser = user;

      loadingEl.style.display = 'none';
      this.updateHeader();
      this.renderData();
    } catch (error: unknown) {
      loadingEl.style.display = 'none';
      errorEl.style.display = 'block';

      let errorMessage = 'Error: ';
      if (error instanceof Error) {
        errorMessage += error.message;

        if (error.message.includes('403')) {
          errorMessage +=
            '\n\nAccess denied. You may not have permission to view site permissions.';
        }
      } else {
        errorMessage += String(error);
      }

      errorEl.textContent = errorMessage;
      console.error('Permission Inspector error:', error);
    }
  }

  /**
   * Update header
   */
  private updateHeader(): void {
    const titleEl = document.getElementById('sp-perm-inspector-title');
    const subtitleEl = document.getElementById('sp-perm-inspector-subtitle');

    if (!titleEl || !subtitleEl) return;

    titleEl.textContent = 'Permission Inspector';

    if (this.permissionInfo) {
      subtitleEl.style.display = 'flex';
      const uniqueText = this.permissionInfo.hasUniqueRoleAssignments
        ? 'Unique permissions'
        : 'Inheriting permissions';
      subtitleEl.innerHTML = `
        <span>${this.escapeHtml(this.permissionInfo.scopeTitle || 'Site')}</span>
        <span style="color: #999;">|</span>
        <span>${uniqueText}</span>
      `;
    }
  }

  /**
   * Render data based on current view mode
   */
  private renderData(): void {
    const dataEl = document.getElementById('sp-perm-inspector-data');
    if (!dataEl) return;

    dataEl.style.display = 'block';

    switch (this.viewMode) {
      case 'assignments':
        this.renderRoleAssignments(dataEl);
        break;
      case 'groups':
        this.renderGroups(dataEl);
        break;
      case 'levels':
        this.renderPermissionLevels(dataEl);
        break;
    }
  }

  /**
   * Render role assignments view
   */
  private renderRoleAssignments(container: HTMLElement): void {
    if (!this.permissionInfo) {
      container.innerHTML = '<div class="sp-loading">No permission data</div>';
      return;
    }

    let assignments = this.permissionInfo.roleAssignments;

    // Filter
    if (this.filterText) {
      const search = this.filterText.toLowerCase();
      assignments = assignments.filter(
        (ra) =>
          ra.principalTitle.toLowerCase().includes(search) ||
          ra.roleDefinitions.some((rd) => rd.name.toLowerCase().includes(search))
      );
    }

    if (this.selectedIndex >= assignments.length) {
      this.selectedIndex = Math.max(0, assignments.length - 1);
    }

    const context = getSharePointContext();
    const webUrl = context?.webUrl || '';

    let html = `
      <div class="sp-perm-info-banner">
        <strong>${assignments.length}</strong> role assignment(s)
        ${webUrl ? `<a href="${sanitizeUrl(getSitePermissionsUrl(webUrl))}" target="_blank" rel="noopener noreferrer" class="sp-perm-link-btn">⚙️ Manage Permissions</a>` : ''}
      </div>
    `;

    html += assignments
      .map((ra, index) => {
        const isSelected = index === this.selectedIndex;
        const typeClass = ra.principalType === 'user' ? 'user' : 'group';

        return `
          <div class="sp-perm-card ${isSelected ? 'selected' : ''}" data-index="${index}">
            <div class="sp-perm-card-header">
              <div>
                <div class="sp-perm-card-title">
                  ${ra.principalType === 'user' ? '👤' : '👥'} ${this.escapeHtml(ra.principalTitle)}
                </div>
                ${ra.principalEmail ? `<div class="sp-perm-card-subtitle">${this.escapeHtml(ra.principalEmail)}</div>` : ''}
              </div>
              <span class="sp-perm-type-badge ${typeClass}">${ra.principalType}</span>
            </div>
            <div class="sp-perm-roles">
              ${ra.roleDefinitions.map((rd) => `<span class="sp-perm-role-chip">${this.escapeHtml(rd.name)}</span>`).join('')}
            </div>
          </div>
        `;
      })
      .join('');

    container.innerHTML = html;
  }

  /**
   * Render groups view
   */
  private renderGroups(container: HTMLElement): void {
    let groups = this.siteGroups;

    if (this.filterText) {
      const search = this.filterText.toLowerCase();
      groups = groups.filter(
        (g) =>
          g.title.toLowerCase().includes(search) || g.description.toLowerCase().includes(search)
      );
    }

    if (this.selectedIndex >= groups.length) {
      this.selectedIndex = Math.max(0, groups.length - 1);
    }

    const context = getSharePointContext();
    const webUrl = context?.webUrl || '';

    let html = `
      <div class="sp-perm-info-banner">
        <strong>${groups.length}</strong> SharePoint group(s)
      </div>
    `;

    html += groups
      .map((group, index) => {
        const isSelected = index === this.selectedIndex;
        const isExpanded = this.expandedGroups.has(group.id);
        const users = this.groupUsers.get(group.id);

        let usersHtml = '';
        if (isExpanded) {
          if (users) {
            usersHtml = `
              <div class="sp-perm-group-users">
                <div class="sp-perm-group-users-header">
                  <span class="sp-perm-group-users-title">Members (${users.length})</span>
                </div>
                <div class="sp-perm-user-list">
                  ${users
                    .map(
                      (u) => `
                    <div class="sp-perm-user-item">
                      <span class="sp-perm-user-icon">${u.isSiteAdmin ? '⭐' : '👤'}</span>
                      <span class="sp-perm-user-name">${this.escapeHtml(u.title)}</span>
                      ${u.email ? `<span class="sp-perm-user-email">${this.escapeHtml(u.email)}</span>` : ''}
                    </div>
                  `
                    )
                    .join('')}
                </div>
              </div>
            `;
          } else {
            usersHtml = `<div class="sp-perm-group-users"><div class="sp-loading">Loading members...</div></div>`;
          }
        }

        return `
          <div class="sp-perm-card ${isSelected ? 'selected' : ''}" data-index="${index}" data-group-id="${group.id}">
            <div class="sp-perm-card-header">
              <div>
                <div class="sp-perm-card-title">👥 ${this.escapeHtml(group.title)}</div>
                ${group.description ? `<div class="sp-perm-card-subtitle">${this.escapeHtml(group.description)}</div>` : ''}
              </div>
              <button class="sp-perm-expand-btn" data-group-id="${group.id}">
                ${isExpanded ? '▼ Hide' : '▶ Show'} Members
              </button>
            </div>
            <div class="sp-perm-details">
              <span class="sp-perm-detail-label">Owner:</span>
              <span class="sp-perm-detail-value">${this.escapeHtml(group.ownerTitle)}</span>
            </div>
            ${webUrl ? `<a href="${sanitizeUrl(getGroupManagementUrl(webUrl, group.id))}" target="_blank" rel="noopener noreferrer" class="sp-perm-link-btn">⚙️ Manage Group</a>` : ''}
            ${usersHtml}
          </div>
        `;
      })
      .join('');

    container.innerHTML = html;
  }

  /**
   * Render permission levels view
   */
  private renderPermissionLevels(container: HTMLElement): void {
    let levels = this.permissionLevels;

    if (this.filterText) {
      const search = this.filterText.toLowerCase();
      levels = levels.filter(
        (l) => l.name.toLowerCase().includes(search) || l.description.toLowerCase().includes(search)
      );
    }

    if (this.selectedIndex >= levels.length) {
      this.selectedIndex = Math.max(0, levels.length - 1);
    }

    let html = `
      <div class="sp-perm-info-banner">
        <strong>${levels.length}</strong> permission level(s)
      </div>
    `;

    html += levels
      .map((level, index) => {
        const isSelected = index === this.selectedIndex;
        const roleTypeName = getRoleTypeName(level.roleTypeKind);

        return `
          <div class="sp-perm-card ${isSelected ? 'selected' : ''}" data-index="${index}">
            <div class="sp-perm-card-header">
              <div>
                <div class="sp-perm-card-title">🔐 ${this.escapeHtml(level.name)}</div>
                ${level.description ? `<div class="sp-perm-card-subtitle">${this.escapeHtml(level.description)}</div>` : ''}
              </div>
              ${level.hidden ? '<span class="sp-perm-type-badge hidden">Hidden</span>' : ''}
            </div>
            <div class="sp-perm-details">
              <span class="sp-perm-detail-label">Type:</span>
              <span class="sp-perm-detail-value">${roleTypeName}</span>
              <span class="sp-perm-detail-label">ID:</span>
              <span class="sp-perm-detail-value">${level.id}</span>
            </div>
          </div>
        `;
      })
      .join('');

    container.innerHTML = html;
  }

  /**
   * Toggle group expansion to show/hide users
   */
  private async toggleGroupExpansion(groupId: number): Promise<void> {
    if (this.expandedGroups.has(groupId)) {
      this.expandedGroups.delete(groupId);
    } else {
      this.expandedGroups.add(groupId);

      // Load users if not already loaded
      if (!this.groupUsers.has(groupId)) {
        this.renderData(); // Show loading state
        try {
          const context = getSharePointContext();
          let webUrl = context?.webUrl || '';
          const siteUrl = context?.siteUrl || '';
          if (siteUrl.includes('/sites/') && !webUrl.includes('/sites/')) {
            webUrl = siteUrl;
          }
          if (!webUrl) {
            webUrl = window.location.origin;
          }

          const users = await getGroupUsers(webUrl, groupId);
          this.groupUsers.set(groupId, users);
        } catch (error) {
          console.error('Failed to load group users:', error);
          this.groupUsers.set(groupId, []);
        }
      }
    }

    this.renderData();
  }

  /**
   * Scroll selected card into view
   */
  private scrollToSelectedCard(): void {
    if (this.selectedIndex < 0) return;

    const dataEl = document.getElementById('sp-perm-inspector-data');
    if (!dataEl) return;

    const selectedCard = dataEl.querySelector(
      `.sp-perm-card[data-index="${this.selectedIndex}"]`
    ) as HTMLElement;

    if (selectedCard) {
      selectedCard.scrollIntoView({
        behavior: 'smooth',
        block: 'nearest',
      });
    }
  }

  /**
   * Get focusable elements within a card
   */
  private getFocusableElements(card: HTMLElement): HTMLElement[] {
    return Array.from(
      card.querySelectorAll('a.sp-perm-link-btn, button.sp-perm-expand-btn')
    ) as HTMLElement[];
  }

  /**
   * Enter inner navigation mode for a card
   */
  private enterInnerNavigationMode(card: HTMLElement): void {
    const focusable = this.getFocusableElements(card);
    if (focusable.length === 0) return;

    this.innerNavigationMode = true;
    this.innerFocusIndex = 0;
    focusable[0].focus();
    card.classList.add('inner-nav-active');
  }

  /**
   * Exit inner navigation mode
   */
  private exitInnerNavigationMode(): void {
    this.innerNavigationMode = false;
    this.innerFocusIndex = -1;

    const activeCard = document.querySelector('.sp-perm-card.inner-nav-active');
    if (activeCard) {
      activeCard.classList.remove('inner-nav-active');
    }

    this.renderData();
    this.scrollToSelectedCard();
  }

  /**
   * Handle keyboard navigation within a card
   */
  private handleInnerNavigation(e: KeyboardEvent, cards: Element[]): boolean {
    const selectedCard = cards[this.selectedIndex] as HTMLElement;
    if (!selectedCard) return false;

    const focusable = this.getFocusableElements(selectedCard);
    if (focusable.length === 0) return false;

    switch (e.key) {
      case 'ArrowDown':
        this.innerFocusIndex = Math.min(this.innerFocusIndex + 1, focusable.length - 1);
        focusable[this.innerFocusIndex].focus();
        return true;

      case 'ArrowUp':
        this.innerFocusIndex = Math.max(this.innerFocusIndex - 1, 0);
        focusable[this.innerFocusIndex].focus();
        return true;

      case 'Enter':
        focusable[this.innerFocusIndex]?.click();
        return true;

      case 'Escape':
        this.exitInnerNavigationMode();
        return true;

      case 'Tab':
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
   * Escape HTML to prevent XSS
   */
  private escapeHtml(text: string): string {
    return escapeHtmlValue(text);
  }
}
