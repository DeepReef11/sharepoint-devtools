/**
 * SharePoint DevTools - Options Page Controller
 * Manages the custom links settings UI
 */

import { CustomLinksStorage } from '../storage/custom-links-storage';
import { escapeHtml as escapeHtmlValue } from '../utils/html';
import { SharePointLink, LinkCategory } from '../links/types';
import { LinkTemplateEngine } from '../links/template-engine';

/**
 * Options page controller
 */
class OptionsController {
  private links: SharePointLink[] = [];
  private editingLinkId: string | null = null;

  /**
   * Initialize the options page
   */
  async init(): Promise<void> {
    await this.loadLinks();
    this.setupEventListeners();
    this.updateUI();
  }

  /**
   * Load custom links from storage
   */
  async loadLinks(): Promise<void> {
    const result = await CustomLinksStorage.loadCustomLinks();
    if (result.success && result.data) {
      this.links = result.data;
    } else {
      this.showNotification(`Failed to load links: ${result.error}`, 'error');
      this.links = [];
    }
  }

  /**
   * Setup event listeners
   */
  setupEventListeners(): void {
    // Action buttons
    document.getElementById('addLinkBtn')?.addEventListener('click', () => this.openAddModal());
    document.getElementById('importBtn')?.addEventListener('click', () => this.openImportModal());
    document.getElementById('exportBtn')?.addEventListener('click', () => this.exportLinks());
    document.getElementById('clearAllBtn')?.addEventListener('click', () => this.clearAllLinks());

    // Link form
    document.getElementById('linkForm')?.addEventListener('submit', (e) => {
      e.preventDefault();
      this.saveLink();
    });
    document.getElementById('cancelBtn')?.addEventListener('click', () => this.closeModal());
    document.getElementById('testLinkBtn')?.addEventListener('click', () => this.testLink());

    // Modal close buttons
    document.getElementById('closeModal')?.addEventListener('click', () => this.closeModal());
    document
      .getElementById('closeImportModal')
      ?.addEventListener('click', () => this.closeImportModal());

    // Import modal
    document
      .getElementById('cancelImportBtn')
      ?.addEventListener('click', () => this.closeImportModal());
    document
      .getElementById('executeImportBtn')
      ?.addEventListener('click', () => this.executeImport());
    document
      .getElementById('importFile')
      ?.addEventListener('change', (e) => this.handleFileSelect(e as Event));

    // Close modals on backdrop click
    document.getElementById('linkModal')?.addEventListener('click', (e) => {
      if ((e.target as HTMLElement).id === 'linkModal') {
        this.closeModal();
      }
    });
    document.getElementById('importModal')?.addEventListener('click', (e) => {
      if ((e.target as HTMLElement).id === 'importModal') {
        this.closeImportModal();
      }
    });
  }

  /**
   * Update UI elements
   */
  async updateUI(): Promise<void> {
    this.renderLinksList();
    await this.updateStats();
  }

  /**
   * Render links list
   */
  renderLinksList(): void {
    const container = document.getElementById('linksContainer');
    const emptyState = document.getElementById('emptyState');

    if (!container) return;

    if (this.links.length === 0) {
      if (emptyState) {
        emptyState.style.display = 'block';
      }
      // Clear any existing link cards
      const linkCards = container.querySelectorAll('.link-card');
      linkCards.forEach((card) => card.remove());
      return;
    }

    if (emptyState) {
      emptyState.style.display = 'none';
    }

    // Clear existing cards
    const linkCards = container.querySelectorAll('.link-card');
    linkCards.forEach((card) => card.remove());

    // Render each link
    this.links.forEach((link) => {
      const card = this.createLinkCard(link);
      container.appendChild(card);
    });
  }

  /**
   * Create a link card element
   */
  createLinkCard(link: SharePointLink): HTMLElement {
    const card = document.createElement('div');
    card.className = 'link-card';
    card.dataset.linkId = link.id;

    const contextBadges = this.getContextBadges(link);
    const keywords = link.keywords ? link.keywords.join(', ') : 'None';
    const priority = link.priority !== undefined ? link.priority : 'Not set';

    card.innerHTML = `
      <div class="link-card-header">
        <div>
          <div class="link-card-title">${this.escapeHtml(link.title)}</div>
          <span class="link-card-category">${this.escapeHtml(link.category)}</span>
        </div>
      </div>
      <div class="link-card-description">${this.escapeHtml(link.description)}</div>
      <div class="link-card-url">${this.escapeHtml(link.urlTemplate)}</div>
      <div class="link-card-meta">
        <div class="link-card-meta-item">
          <strong>ID:</strong> ${this.escapeHtml(link.id)}
        </div>
        <div class="link-card-meta-item">
          <strong>Priority:</strong> ${priority}
        </div>
        ${contextBadges ? `<div class="link-card-meta-item">${contextBadges}</div>` : ''}
      </div>
      ${
        link.keywords && link.keywords.length > 0
          ? `
      <div class="link-card-meta">
        <div class="link-card-meta-item">
          <strong>Keywords:</strong> ${keywords}
        </div>
      </div>
      `
          : ''
      }
      <div class="link-card-actions">
        <button class="btn btn-secondary btn-small edit-btn" data-link-id="${link.id}">
          Edit
        </button>
        <button class="btn btn-danger btn-small delete-btn" data-link-id="${link.id}">
          Delete
        </button>
      </div>
    `;

    // Attach event listeners
    card.querySelector('.edit-btn')?.addEventListener('click', () => this.editLink(link.id));
    card.querySelector('.delete-btn')?.addEventListener('click', () => this.deleteLink(link.id));

    return card;
  }

  /**
   * Get context requirement badges
   */
  getContextBadges(link: SharePointLink): string {
    if (!link.context) return '';

    const badges: string[] = [];

    if (link.context.requiresList) badges.push('List Context');
    if (link.context.requiresSiteAdmin) badges.push('Site Admin');
    if (link.context.requiresTenantAdmin) badges.push('Tenant Admin');
    if (link.context.modernOnly) badges.push('Modern Only');
    if (link.context.classicOnly) badges.push('Classic Only');

    if (badges.length === 0) return '';

    return `<strong>Context:</strong> ${badges.join(', ')}`;
  }

  /**
   * Update statistics
   */
  async updateStats(): Promise<void> {
    const linkCountEl = document.getElementById('linkCount');
    const storageUsedEl = document.getElementById('storageUsed');

    if (linkCountEl) {
      linkCountEl.textContent = this.links.length.toString();
    }

    const statsResult = await CustomLinksStorage.getStorageStats();
    if (statsResult.success && statsResult.data && storageUsedEl) {
      const kb = (statsResult.data.bytesUsed / 1024).toFixed(2);
      const quotaKb = (statsResult.data.quota / 1024).toFixed(0);
      storageUsedEl.textContent = `${kb} KB / ${quotaKb} KB`;
    }
  }

  /**
   * Open add new link modal
   */
  openAddModal(): void {
    this.editingLinkId = null;
    this.resetForm();
    document.getElementById('modalTitle')!.textContent = 'Add New Link';
    document.getElementById('linkModal')?.classList.add('active');
    document.getElementById('linkId')?.removeAttribute('readonly');
  }

  /**
   * Edit existing link
   */
  editLink(linkId: string): void {
    const link = this.links.find((l) => l.id === linkId);
    if (!link) {
      this.showNotification('Link not found', 'error');
      return;
    }

    this.editingLinkId = linkId;
    this.populateForm(link);
    document.getElementById('modalTitle')!.textContent = 'Edit Link';
    document.getElementById('linkModal')?.classList.add('active');
    document.getElementById('linkId')?.setAttribute('readonly', 'readonly');
  }

  /**
   * Populate form with link data
   */
  populateForm(link: SharePointLink): void {
    (document.getElementById('linkId') as HTMLInputElement).value = link.id;
    (document.getElementById('linkTitle') as HTMLInputElement).value = link.title;
    (document.getElementById('linkUrlTemplate') as HTMLInputElement).value = link.urlTemplate;
    (document.getElementById('linkCategory') as HTMLSelectElement).value = link.category;
    (document.getElementById('linkDescription') as HTMLTextAreaElement).value = link.description;
    (document.getElementById('linkKeywords') as HTMLInputElement).value =
      link.keywords?.join(', ') || '';
    (document.getElementById('linkPriority') as HTMLInputElement).value =
      link.priority?.toString() || '';

    // Context checkboxes
    (document.getElementById('requiresList') as HTMLInputElement).checked =
      link.context?.requiresList || false;
    (document.getElementById('requiresSiteAdmin') as HTMLInputElement).checked =
      link.context?.requiresSiteAdmin || false;
    (document.getElementById('requiresTenantAdmin') as HTMLInputElement).checked =
      link.context?.requiresTenantAdmin || false;
    (document.getElementById('modernOnly') as HTMLInputElement).checked =
      link.context?.modernOnly || false;
    (document.getElementById('classicOnly') as HTMLInputElement).checked =
      link.context?.classicOnly || false;
  }

  /**
   * Reset form
   */
  resetForm(): void {
    (document.getElementById('linkForm') as HTMLFormElement).reset();
  }

  /**
   * Save link (add or update)
   */
  async saveLink(): Promise<void> {
    const link = this.getLinkFromForm();
    if (!link) return;

    // Validate link
    const validation = this.validateLink(link);
    if (!validation.valid) {
      this.showNotification(validation.error || 'Invalid link data', 'error');
      return;
    }

    let result;
    if (this.editingLinkId) {
      // Update existing link
      result = await CustomLinksStorage.updateCustomLink(this.editingLinkId, link);
    } else {
      // Add new link
      result = await CustomLinksStorage.addCustomLink(link);
    }

    if (result.success) {
      this.showNotification(
        this.editingLinkId ? 'Link updated successfully' : 'Link added successfully',
        'success'
      );
      await this.loadLinks();
      this.updateUI();
      this.closeModal();
    } else {
      this.showNotification(`Failed to save link: ${result.error}`, 'error');
    }
  }

  /**
   * Get link data from form
   */
  getLinkFromForm(): SharePointLink | null {
    const id = (document.getElementById('linkId') as HTMLInputElement).value.trim();
    const title = (document.getElementById('linkTitle') as HTMLInputElement).value.trim();
    const urlTemplate = (
      document.getElementById('linkUrlTemplate') as HTMLInputElement
    ).value.trim();
    const category = (document.getElementById('linkCategory') as HTMLSelectElement)
      .value as LinkCategory;
    const description = (
      document.getElementById('linkDescription') as HTMLTextAreaElement
    ).value.trim();
    const keywordsStr = (document.getElementById('linkKeywords') as HTMLInputElement).value.trim();
    const priorityStr = (document.getElementById('linkPriority') as HTMLInputElement).value.trim();

    const keywords = keywordsStr
      ? keywordsStr
          .split(',')
          .map((k) => k.trim())
          .filter((k) => k)
      : undefined;
    const priority = priorityStr ? parseInt(priorityStr, 10) : undefined;

    const requiresList = (document.getElementById('requiresList') as HTMLInputElement).checked;
    const requiresSiteAdmin = (document.getElementById('requiresSiteAdmin') as HTMLInputElement)
      .checked;
    const requiresTenantAdmin = (document.getElementById('requiresTenantAdmin') as HTMLInputElement)
      .checked;
    const modernOnly = (document.getElementById('modernOnly') as HTMLInputElement).checked;
    const classicOnly = (document.getElementById('classicOnly') as HTMLInputElement).checked;

    const context =
      requiresList || requiresSiteAdmin || requiresTenantAdmin || modernOnly || classicOnly
        ? {
            requiresList: requiresList || undefined,
            requiresSiteAdmin: requiresSiteAdmin || undefined,
            requiresTenantAdmin: requiresTenantAdmin || undefined,
            modernOnly: modernOnly || undefined,
            classicOnly: classicOnly || undefined,
          }
        : undefined;

    return {
      id,
      title,
      urlTemplate,
      category,
      description,
      keywords,
      priority,
      context,
    };
  }

  /**
   * Validate link data
   */
  validateLink(link: SharePointLink): { valid: boolean; error?: string } {
    if (!link.id || !/^[a-z0-9-]+$/.test(link.id)) {
      return {
        valid: false,
        error: 'Invalid ID. Use lowercase letters, numbers, and hyphens only.',
      };
    }

    if (!link.title) {
      return { valid: false, error: 'Title is required.' };
    }

    if (!link.urlTemplate) {
      return { valid: false, error: 'URL template is required.' };
    }

    if (!link.category) {
      return { valid: false, error: 'Category is required.' };
    }

    if (!link.description) {
      return { valid: false, error: 'Description is required.' };
    }

    // Check for conflicting context
    if (link.context?.modernOnly && link.context?.classicOnly) {
      return { valid: false, error: 'Cannot require both Modern and Classic SharePoint.' };
    }

    return { valid: true };
  }

  /**
   * Test link URL
   */
  testLink(): void {
    const link = this.getLinkFromForm();
    if (!link) return;

    // Create test placeholder values
    const testValues = {
      webUrl: 'https://contoso.sharepoint.com/sites/mysite',
      siteUrl: 'https://contoso.sharepoint.com/sites/mysite',
      listId: '{12345678-1234-1234-1234-123456789012}',
      listUrl: 'https://contoso.sharepoint.com/sites/mysite/Lists/MyList',
    };

    const resolvedUrl = LinkTemplateEngine.resolve(link.urlTemplate, testValues);

    if (resolvedUrl) {
      this.showNotification(`Test URL: ${resolvedUrl}`, 'success');
    } else {
      this.showNotification('Failed to resolve URL template with test values', 'error');
    }
  }

  /**
   * Delete link
   */
  async deleteLink(linkId: string): Promise<void> {
    if (!confirm('Are you sure you want to delete this link?')) {
      return;
    }

    const result = await CustomLinksStorage.deleteCustomLink(linkId);

    if (result.success) {
      this.showNotification('Link deleted successfully', 'success');
      await this.loadLinks();
      this.updateUI();
    } else {
      this.showNotification(`Failed to delete link: ${result.error}`, 'error');
    }
  }

  /**
   * Close link modal
   */
  closeModal(): void {
    document.getElementById('linkModal')?.classList.remove('active');
    this.editingLinkId = null;
    this.resetForm();
  }

  /**
   * Export links
   */
  async exportLinks(): Promise<void> {
    const result = await CustomLinksStorage.exportCustomLinks();

    if (result.success && result.data) {
      const blob = new Blob([result.data], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `sharepoint-devtools-links-${new Date().toISOString().split('T')[0]}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      this.showNotification('Links exported successfully', 'success');
    } else {
      this.showNotification(`Failed to export links: ${result.error}`, 'error');
    }
  }

  /**
   * Open import modal
   */
  openImportModal(): void {
    document.getElementById('importModal')?.classList.add('active');
  }

  /**
   * Close import modal
   */
  closeImportModal(): void {
    document.getElementById('importModal')?.classList.remove('active');
    (document.getElementById('importFile') as HTMLInputElement).value = '';
    (document.getElementById('importJson') as HTMLTextAreaElement).value = '';
  }

  /**
   * Handle file selection
   */
  handleFileSelect(event: Event): void {
    const target = event.target as HTMLInputElement;
    const file = target.files?.[0];

    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const content = e.target?.result as string;
      (document.getElementById('importJson') as HTMLTextAreaElement).value = content;
    };
    reader.readAsText(file);
  }

  /**
   * Execute import
   */
  async executeImport(): Promise<void> {
    const jsonData = (document.getElementById('importJson') as HTMLTextAreaElement).value.trim();

    if (!jsonData) {
      this.showNotification('Please provide JSON data to import', 'error');
      return;
    }

    const mode = (document.querySelector('input[name="importMode"]:checked') as HTMLInputElement)
      ?.value as 'replace' | 'merge';

    const result = await CustomLinksStorage.importCustomLinks(jsonData, mode);

    if (result.success) {
      this.showNotification(`Links imported successfully (${mode} mode)`, 'success');
      await this.loadLinks();
      this.updateUI();
      this.closeImportModal();
    } else {
      this.showNotification(`Import failed: ${result.error}`, 'error');
    }
  }

  /**
   * Clear all links
   */
  async clearAllLinks(): Promise<void> {
    if (!confirm('Are you sure you want to delete ALL custom links? This cannot be undone.')) {
      return;
    }

    const result = await CustomLinksStorage.clearCustomLinks();

    if (result.success) {
      this.showNotification('All custom links cleared', 'success');
      await this.loadLinks();
      this.updateUI();
    } else {
      this.showNotification(`Failed to clear links: ${result.error}`, 'error');
    }
  }

  /**
   * Show notification
   */
  showNotification(message: string, type: 'success' | 'error' = 'success'): void {
    const notification = document.getElementById('notification');
    if (!notification) return;

    notification.textContent = message;
    notification.className = `notification ${type} show`;

    setTimeout(() => {
      notification.classList.remove('show');
    }, 5000);
  }

  /**
   * Escape HTML to prevent XSS
   */
  escapeHtml(text: string): string {
    return escapeHtmlValue(text);
  }
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    const controller = new OptionsController();
    controller.init();
  });
} else {
  const controller = new OptionsController();
  controller.init();
}
