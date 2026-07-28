/**
 * SharePoint Object Inspector Component
 * Displays SharePoint metadata in a readable format with copy-to-clipboard functionality
 */

/* eslint-disable @typescript-eslint/no-explicit-any */

import {
  getContextMetadata,
  getTemplateTypeName,
  getFieldTypeDisplayName,
  type ListMetadata,
  type FieldMetadata,
  type ContentTypeMetadata,
  type WebMetadata,
  type SiteMetadata,
} from '../api/sharepoint-api';
import {
  getSharePointContext,
  extractListIdAsync,
  isSharePointPage,
} from '../context/sharepoint-context';
import { escapeHtml as escapeHtmlValue } from '../utils/html';

export class ObjectInspector {
  private container: HTMLElement | null = null;
  private isLoading = false;

  constructor(containerElement: HTMLElement) {
    this.container = containerElement;
  }

  /**
   * Initializes and loads the inspector with current context
   */
  public async load(): Promise<void> {
    if (!this.container) {
      return;
    }

    this.showLoading();

    try {
      // Check if this is a SharePoint page using the updated detection
      if (!isSharePointPage()) {
        this.showError('Not a valid SharePoint page');
        return;
      }

      // Get the basic context
      const context = getSharePointContext();

      if (!context.webUrl) {
        this.showError('Unable to detect SharePoint web URL');
        return;
      }

      // For classic SharePoint pages, try to get list ID asynchronously
      let listId = context.listId;
      if (!listId) {
        console.log('Object Inspector - No list ID in context, trying async extraction');
        listId = await extractListIdAsync();
        if (listId) {
          console.log('Object Inspector - Got list ID from async extraction:', listId);
          // Update context with the fetched list ID
          context.listId = listId;
        }
      }

      // Fetch metadata from SharePoint API
      const metadata = await getContextMetadata(context.webUrl, listId);

      if (metadata.error) {
        this.showError(metadata.error);
        return;
      }

      this.render(metadata, context);
    } catch (error) {
      console.error('Error loading object inspector:', error);
      this.showError(error instanceof Error ? error.message : 'Unknown error occurred');
    }
  }

  /**
   * Shows loading state
   */
  private showLoading(): void {
    if (!this.container) return;

    this.isLoading = true;
    this.container.innerHTML = `
      <div class="sp-inspector-loading">
        <div class="sp-spinner"></div>
        <p>Loading SharePoint metadata...</p>
      </div>
    `;
  }

  /**
   * Shows error state
   */
  private showError(message: string): void {
    if (!this.container) return;

    this.isLoading = false;
    this.container.innerHTML = `
      <div class="sp-inspector-error">
        <p class="sp-error-icon">⚠️</p>
        <p class="sp-error-message">${this.escapeHtml(message)}</p>
      </div>
    `;
  }

  /**
   * Renders the metadata
   */
  private render(metadata: any, context: any): void {
    if (!this.container) return;

    this.isLoading = false;
    this.container.innerHTML = '';

    const wrapper = document.createElement('div');
    wrapper.className = 'sp-inspector-wrapper';

    // Add title
    const title = document.createElement('h3');
    title.className = 'sp-inspector-title';
    title.textContent = 'SharePoint Object Inspector';
    wrapper.appendChild(title);

    // Add context info
    const contextInfo = document.createElement('div');
    contextInfo.className = 'sp-inspector-context';
    contextInfo.innerHTML = `
      <span class="sp-context-badge">${context.isModern ? 'Modern' : 'Classic'} UI</span>
      <span class="sp-context-badge">${context.pageType || 'unknown'}</span>
    `;
    wrapper.appendChild(contextInfo);

    // Create tabs for different sections
    const tabs = this.createTabs(metadata, context);
    wrapper.appendChild(tabs);

    this.container.appendChild(wrapper);
  }

  /**
   * Creates tabbed interface for different metadata sections
   */
  private createTabs(metadata: any, context: any): HTMLElement {
    const tabsContainer = document.createElement('div');
    tabsContainer.className = 'sp-inspector-tabs';

    const tabButtons = document.createElement('div');
    tabButtons.className = 'sp-tab-buttons';

    const tabContents = document.createElement('div');
    tabContents.className = 'sp-tab-contents';

    const tabs: { id: string; label: string; content: HTMLElement }[] = [];

    // Web/Site tab
    if (metadata.web || metadata.site) {
      tabs.push({
        id: 'web',
        label: 'Web/Site',
        content: this.createWebSection(metadata.web, metadata.site, context),
      });
    }

    // List/Library tab
    if (metadata.list) {
      tabs.push({
        id: 'list',
        label: 'List/Library',
        content: this.createListSection(metadata.list),
      });
    }

    // Fields tab
    if (metadata.fields && metadata.fields.length > 0) {
      tabs.push({
        id: 'fields',
        label: `Fields (${metadata.fields.length})`,
        content: this.createFieldsSection(metadata.fields),
      });
    }

    // Content Types tab
    if (metadata.contentTypes && metadata.contentTypes.length > 0) {
      tabs.push({
        id: 'contentTypes',
        label: `Content Types (${metadata.contentTypes.length})`,
        content: this.createContentTypesSection(metadata.contentTypes),
      });
    }

    // Create tab buttons and contents
    tabs.forEach((tab, index) => {
      const button = document.createElement('button');
      button.className = 'sp-tab-button';
      button.textContent = tab.label;
      button.dataset.tab = tab.id;
      if (index === 0) {
        button.classList.add('active');
      }

      button.onclick = () => {
        // Remove active class from all buttons and contents
        tabButtons.querySelectorAll('.sp-tab-button').forEach((b) => b.classList.remove('active'));
        tabContents
          .querySelectorAll('.sp-tab-content')
          .forEach((c) => c.classList.remove('active'));

        // Add active class to clicked button and corresponding content
        button.classList.add('active');
        tab.content.classList.add('active');
      };

      tabButtons.appendChild(button);

      tab.content.className = 'sp-tab-content';
      tab.content.dataset.tab = tab.id;
      if (index === 0) {
        tab.content.classList.add('active');
      }
      tabContents.appendChild(tab.content);
    });

    tabsContainer.appendChild(tabButtons);
    tabsContainer.appendChild(tabContents);

    return tabsContainer;
  }

  /**
   * Creates the Web/Site metadata section
   */
  private createWebSection(web: WebMetadata, site: SiteMetadata, _context: any): HTMLElement {
    const section = document.createElement('div');
    section.className = 'sp-metadata-section';

    const items: { label: string; value: string; copyable?: boolean }[] = [];

    if (web) {
      items.push(
        { label: 'Web Title', value: web.title },
        { label: 'Web ID', value: web.id, copyable: true },
        { label: 'Web URL', value: web.url, copyable: true },
        { label: 'Server Relative URL', value: web.serverRelativeUrl, copyable: true },
        { label: 'Web Template', value: web.webTemplate },
        { label: 'Language', value: this.getLanguageName(web.language) },
        { label: 'Created', value: new Date(web.created).toLocaleString() }
      );

      if (web.description) {
        items.push({ label: 'Description', value: web.description });
      }
    }

    if (site) {
      items.push(
        { label: 'Site Collection ID', value: site.id, copyable: true },
        { label: 'Site Collection URL', value: site.url, copyable: true }
      );
    }

    section.appendChild(this.createMetadataList(items));

    return section;
  }

  /**
   * Creates the List/Library metadata section
   */
  private createListSection(list: ListMetadata): HTMLElement {
    const section = document.createElement('div');
    section.className = 'sp-metadata-section';

    const items: { label: string; value: string; copyable?: boolean }[] = [
      { label: 'Title', value: list.title },
      { label: 'Internal Name', value: list.internalName, copyable: true },
      { label: 'List ID', value: list.id, copyable: true },
      {
        label: 'Template Type',
        value: `${getTemplateTypeName(list.baseTemplate)} (${list.baseTemplate})`,
      },
      { label: 'Base Type', value: list.baseType === 0 ? 'List' : 'Document Library' },
      { label: 'Item Count', value: list.itemCount.toString() },
      { label: 'Created', value: new Date(list.created).toLocaleString() },
      { label: 'Last Modified', value: new Date(list.lastItemModifiedDate).toLocaleString() },
    ];

    if (list.description) {
      items.push({ label: 'Description', value: list.description });
    }

    section.appendChild(this.createMetadataList(items));

    return section;
  }

  /**
   * Creates the Fields metadata section
   */
  private createFieldsSection(fields: FieldMetadata[]): HTMLElement {
    const section = document.createElement('div');
    section.className = 'sp-metadata-section sp-fields-section';

    // Add search/filter input
    const filterInput = document.createElement('input');
    filterInput.type = 'text';
    filterInput.className = 'sp-field-filter';
    filterInput.placeholder = 'Filter fields...';
    section.appendChild(filterInput);

    const table = document.createElement('table');
    table.className = 'sp-fields-table';

    const thead = document.createElement('thead');
    thead.innerHTML = `
      <tr>
        <th>Title</th>
        <th>Internal Name</th>
        <th>Type</th>
        <th>Required</th>
        <th>Actions</th>
      </tr>
    `;
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    fields.forEach((field) => {
      const row = document.createElement('tr');
      row.className = 'sp-field-row';
      row.dataset.title = field.title.toLowerCase();
      row.dataset.internalName = field.internalName.toLowerCase();

      row.innerHTML = `
        <td class="sp-field-title">${this.escapeHtml(field.title)}</td>
        <td class="sp-field-internal">
          <code>${this.escapeHtml(field.internalName)}</code>
        </td>
        <td class="sp-field-type">
          <span class="sp-type-badge">${this.escapeHtml(getFieldTypeDisplayName(field.typeAsString))}</span>
        </td>
        <td class="sp-field-required">${field.required ? '✓' : ''}</td>
        <td class="sp-field-actions">
          <button class="sp-btn-copy" data-value="${this.escapeHtml(field.internalName)}" title="Copy internal name">
            📋
          </button>
        </td>
      `;

      tbody.appendChild(row);
    });

    table.appendChild(tbody);
    section.appendChild(table);

    // Add filter functionality
    filterInput.addEventListener('input', (e) => {
      const filter = (e.target as HTMLInputElement).value.toLowerCase();
      tbody.querySelectorAll('.sp-field-row').forEach((row) => {
        const title = (row as HTMLElement).dataset.title || '';
        const internalName = (row as HTMLElement).dataset.internalName || '';
        const matches = title.includes(filter) || internalName.includes(filter);
        (row as HTMLElement).style.display = matches ? '' : 'none';
      });
    });

    return section;
  }

  /**
   * Creates the Content Types metadata section
   */
  private createContentTypesSection(contentTypes: ContentTypeMetadata[]): HTMLElement {
    const section = document.createElement('div');
    section.className = 'sp-metadata-section';

    const list = document.createElement('div');
    list.className = 'sp-content-types-list';

    contentTypes.forEach((ct) => {
      const item = document.createElement('div');
      item.className = 'sp-content-type-item';

      item.innerHTML = `
        <div class="sp-ct-header">
          <h4 class="sp-ct-name">${this.escapeHtml(ct.name)}</h4>
          <button class="sp-btn-copy" data-value="${this.escapeHtml(ct.id)}" title="Copy content type ID">
            📋
          </button>
        </div>
        <div class="sp-ct-details">
          <p><strong>ID:</strong> <code>${this.escapeHtml(ct.id)}</code></p>
          ${ct.description ? `<p><strong>Description:</strong> ${this.escapeHtml(ct.description)}</p>` : ''}
          ${ct.group ? `<p><strong>Group:</strong> ${this.escapeHtml(ct.group)}</p>` : ''}
        </div>
      `;

      list.appendChild(item);
    });

    section.appendChild(list);

    return section;
  }

  /**
   * Creates a metadata list from items
   */
  private createMetadataList(
    items: { label: string; value: string; copyable?: boolean }[]
  ): HTMLElement {
    const list = document.createElement('dl');
    list.className = 'sp-metadata-list';

    items.forEach((item) => {
      const dt = document.createElement('dt');
      dt.textContent = item.label;
      list.appendChild(dt);

      const dd = document.createElement('dd');

      if (item.copyable) {
        dd.innerHTML = `
          <span class="sp-metadata-value">${this.escapeHtml(item.value)}</span>
          <button class="sp-btn-copy" data-value="${this.escapeHtml(item.value)}" title="Copy to clipboard">
            📋
          </button>
        `;
      } else {
        dd.innerHTML = `<span class="sp-metadata-value">${this.escapeHtml(item.value)}</span>`;
      }

      list.appendChild(dd);
    });

    return list;
  }

  /**
   * Gets language name from LCID
   */
  private getLanguageName(lcid: number): string {
    const languages: { [key: number]: string } = {
      1033: 'English (United States)',
      1031: 'German (Germany)',
      1036: 'French (France)',
      1034: 'Spanish (Spain)',
      1040: 'Italian (Italy)',
      1041: 'Japanese',
      2052: 'Chinese (China)',
      1028: 'Chinese (Taiwan)',
      1042: 'Korean',
      1043: 'Dutch (Netherlands)',
      1046: 'Portuguese (Brazil)',
      2070: 'Portuguese (Portugal)',
      1049: 'Russian',
    };

    return languages[lcid] || `Language ${lcid}`;
  }

  /**
   * Escapes HTML to prevent XSS
   */
  private escapeHtml(text: string): string {
    return escapeHtmlValue(text);
  }

  /**
   * Clears the inspector
   */
  public clear(): void {
    if (this.container) {
      this.container.innerHTML = '';
    }
  }
}
