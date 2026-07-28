/**
 * Column Schema Fetcher
 * Fetches column metadata from SharePoint lists/libraries using REST API
 */

import {
  ColumnSchema,
  ListSchema,
  SharePointFieldData,
  ContentTypeInfo,
} from '../types/column-schema';
import { extractListIdAsync } from '../context/sharepoint-context';

export class ColumnFetcher {
  /**
   * Detect if current page is a list/library view
   */
  static isListView(): boolean {
    const url = window.location.href;
    const pathname = window.location.pathname;
    const search = window.location.search.toLowerCase();
    const hash = window.location.hash.toLowerCase();

    // Check for common list view URLs
    const listViewPatterns = [
      '/Lists/',
      '/Forms/AllItems.aspx',
      '/Forms/DispForm.aspx',
      '/Forms/EditForm.aspx',
      '/Forms/NewForm.aspx',
      '/_layouts/15/start.aspx#/Lists/',
      '/_layouts/15/start.aspx#/Shared%20Documents',
      '/Shared%20Documents/',
      '/Shared Documents/',
      '/Documents/',
      '/Forms/', // Generic forms path
    ];

    // Check URL patterns
    if (listViewPatterns.some((pattern) => url.includes(pattern))) {
      return true;
    }

    // Check query parameters (case-insensitive)
    const listParams = ['viewid=', 'rootfolder=', 'list=', 'listid='];
    if (listParams.some((param) => search.includes(param))) {
      return true;
    }

    // Check modern experience hash patterns
    const modernHashPatterns = [
      '#/lists/',
      '#/documents/',
      '#/shared%20documents',
      '#/shared documents',
    ];
    if (modernHashPatterns.some((pattern) => hash.includes(pattern))) {
      return true;
    }

    // Check for document library pattern: /LibraryName/Forms/
    // This catches custom library names
    if (/\/[^/]+\/Forms\/[^/]*\.aspx/i.test(pathname)) {
      return true;
    }

    // Check _spPageContextInfo for list context
    try {
      // @ts-expect-error - SharePoint global context object not in type definitions
      if (window._spPageContextInfo?.listId || window._spPageContextInfo?.pageListId) {
        return true;
      }
    } catch {
      // Ignore errors accessing _spPageContextInfo
    }

    return false;
  }

  /**
   * Extract list ID from current page (synchronous - limited detection)
   * @deprecated Use extractListIdAsync() for better support of classic SharePoint pages
   */
  static extractListId(): string | null {
    // Try to get list ID from URL parameters
    const urlParams = new URLSearchParams(window.location.search);
    const listIdParam = urlParams.get('List');
    if (listIdParam) {
      return listIdParam.replace(/[{}]/g, '');
    }

    // Try to get from page context (if available)
    try {
      // @ts-expect-error - SharePoint global context object not in type definitions
      if (window._spPageContextInfo?.pageListId) {
        // @ts-expect-error - Accessing SharePoint global
        return window._spPageContextInfo.pageListId.replace(/[{}]/g, '');
      }
    } catch (e) {
      console.debug('Could not access _spPageContextInfo:', e);
    }

    // Try to extract from DOM
    const listIdMeta = document.querySelector('meta[name="ListId"]');
    if (listIdMeta) {
      return listIdMeta.getAttribute('content')?.replace(/[{}]/g, '') || null;
    }

    return null;
  }

  /**
   * Extract list ID from current page (async - supports classic SharePoint pages)
   * This method can fetch the list ID from the SharePoint API when it's not
   * immediately available in the page context or URL.
   */
  static async extractListIdAsync(): Promise<string | null> {
    const listId = await extractListIdAsync();
    return listId || null;
  }

  /**
   * Get the current web URL
   */
  static getWebUrl(): string {
    try {
      // @ts-expect-error - SharePoint global context object not in type definitions
      if (window._spPageContextInfo?.webAbsoluteUrl) {
        // @ts-expect-error - Accessing SharePoint global
        return window._spPageContextInfo.webAbsoluteUrl;
      }
    } catch (e) {
      console.debug('Could not access _spPageContextInfo:', e);
    }

    // Fallback: try to construct from URL
    const url = new URL(window.location.href);
    const pathParts = url.pathname.split('/');

    // Look for /sites/ or /teams/ pattern
    const siteIndex = pathParts.findIndex((part) => part === 'sites' || part === 'teams');

    if (siteIndex !== -1 && pathParts[siteIndex + 1]) {
      return `${url.origin}/${pathParts[siteIndex]}/${pathParts[siteIndex + 1]}`;
    }

    // Root site
    return url.origin;
  }

  /**
   * Fetch list schema including all fields
   */
  static async fetchListSchema(listId: string, webUrl?: string): Promise<ListSchema> {
    const baseUrl = webUrl || this.getWebUrl();
    const endpoint = `${baseUrl}/_api/web/lists(guid'${listId}')?$expand=Fields,ContentTypes&$select=*,Fields/*,ContentTypes/*`;

    try {
      const response = await fetch(endpoint, {
        method: 'GET',
        headers: {
          Accept: 'application/json;odata=verbose',
        },
        credentials: 'include',
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();
      const listData = data.d;

      // Parse fields
      const fields: ColumnSchema[] = listData.Fields.results
        .map((field: SharePointFieldData) => this.parseField(field))
        .filter((field: ColumnSchema | null) => field !== null) as ColumnSchema[];

      // Parse content types
      const contentTypes: ContentTypeInfo[] = listData.ContentTypes.results.map(
        (ct: { Id: { StringValue: string }; Name: string; Description?: string }) => ({
          id: ct.Id.StringValue,
          name: ct.Name,
          description: ct.Description || undefined,
        })
      );

      return {
        title: listData.Title,
        internalName: listData.EntityTypeName || listData.Title,
        id: listData.Id,
        baseTemplate: listData.BaseTemplate,
        baseType: this.getBaseTypeName(listData.BaseType),
        description: listData.Description || undefined,
        fields,
        contentTypes,
        contentTypesEnabled: listData.ContentTypesEnabled || false,
        enableVersioning: listData.EnableVersioning,
        hidden: listData.Hidden,
        serverRelativeUrl: listData.RootFolder?.ServerRelativeUrl || '',
        webAbsoluteUrl: baseUrl,
      };
    } catch (error) {
      console.error('Error fetching list schema:', error);
      throw error;
    }
  }

  /**
   * Parse SharePoint field data into ColumnSchema
   */
  private static parseField(field: SharePointFieldData): ColumnSchema | null {
    // Skip system/internal fields that are not useful
    const skipFields = [
      'ContentTypeId',
      'InstanceID',
      '_UIVersionString',
      'owshiddenversion',
      'WorkflowVersion',
      '_Level',
      '_IsCurrentVersion',
      'ItemChildCount',
      'FolderChildCount',
      'AppAuthor',
      'AppEditor',
    ];

    if (skipFields.includes(field.InternalName)) {
      return null;
    }

    const column: ColumnSchema = {
      internalName: field.InternalName,
      title: field.Title,
      type: field.TypeAsString,
      id: field.Id,
      required: field.Required,
      hidden: field.Hidden,
      readOnly: field.ReadOnlyField,
      description: field.Description,
      defaultValue: field.DefaultValue,
      indexed: field.Indexed,
      staticName: field.StaticName,
      group: field.Group,
      maxLength: field.MaxLength,
    };

    // Add type-specific properties
    if (field.TypeAsString === 'Calculated' && field.Formula) {
      column.formula = field.Formula;
    }

    if (field.TypeAsString === 'Lookup' && field.LookupList) {
      column.lookupListId = field.LookupList;
      column.lookupField = field.LookupField;
    }

    if (
      (field.TypeAsString === 'Choice' || field.TypeAsString === 'MultiChoice') &&
      field.Choices
    ) {
      column.choices = field.Choices.results;
    }

    if (field.TypeAsString === 'TaxonomyFieldType' && field.TermSetId) {
      column.termSetId = field.TermSetId;
    }

    return column;
  }

  /**
   * Get base type name from numeric value
   */
  private static getBaseTypeName(baseType: number): string {
    const baseTypes: { [key: number]: string } = {
      0: 'GenericList',
      1: 'DocumentLibrary',
      2: 'Unused',
      3: 'DiscussionBoard',
      4: 'Survey',
      5: 'Issue',
    };

    return baseTypes[baseType] || 'Unknown';
  }

  /**
   * Fetch column details for a specific field
   */
  static async fetchColumnDetails(
    listId: string,
    internalName: string,
    webUrl?: string
  ): Promise<ColumnSchema | null> {
    const baseUrl = webUrl || this.getWebUrl();
    const endpoint = `${baseUrl}/_api/web/lists(guid'${listId}')/fields/getbyinternalnameortitle('${internalName}')`;

    try {
      const response = await fetch(endpoint, {
        method: 'GET',
        headers: {
          Accept: 'application/json;odata=verbose',
        },
        credentials: 'include',
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();
      return this.parseField(data.d);
    } catch (error) {
      console.error('Error fetching column details:', error);
      return null;
    }
  }
}
