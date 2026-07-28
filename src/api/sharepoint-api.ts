/**
 * SharePoint REST API Service
 * Provides methods to fetch SharePoint metadata using REST APIs
 * Now includes caching layer for improved performance
 */

/* eslint-disable @typescript-eslint/no-explicit-any */

import {
  withRetry,
  createErrorFromResponse,
  NetworkError,
  logError,
  logDebug,
} from '../utils/error-handling';
import { apiCache, generateCacheKey, CacheOptions } from './api-cache';

export interface ListMetadata {
  id: string;
  title: string;
  internalName: string;
  baseTemplate: number;
  baseType: number;
  itemCount: number;
  description: string;
  created: string;
  lastItemModifiedDate: string;
  contentTypesEnabled: boolean;
  enableVersioning: boolean;
}

export interface FieldMetadata {
  id: string;
  title: string;
  internalName: string;
  typeAsString: string;
  required: boolean;
  hidden: boolean;
  readOnlyField: boolean;
  description: string;
  group: string;
}

export interface ContentTypeMetadata {
  id: string;
  name: string;
  description: string;
  group: string;
}

export interface SiteMetadata {
  id: string;
  url: string;
  title: string;
  description: string;
  webTemplate: string;
  created: string;
  language: number;
}

export interface WebMetadata {
  id: string;
  title: string;
  description: string;
  serverRelativeUrl: string;
  url: string;
  created: string;
  language: number;
  webTemplate: string;
}

/**
 * Fetches data from SharePoint REST API with retry logic, error handling, and caching support
 */
async function fetchSharePointAPI<T>(url: string, cacheOptions?: CacheOptions): Promise<T> {
  // Check cache first if caching is enabled
  if (cacheOptions) {
    const cacheKey = generateCacheKey(url);
    const cached = apiCache.get<T>(cacheKey, cacheOptions.useSessionStorage);
    if (cached !== null) {
      logDebug('API Cache Hit', { url });
      return cached;
    }
  }

  logDebug('SharePoint API request', { url });

  // Fetch from API with retry logic
  const data = await withRetry(
    async () => {
      let response: Response;

      try {
        response = await fetch(url, {
          method: 'GET',
          headers: {
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
          },
          credentials: 'include',
        });
      } catch (error) {
        // Network errors (no internet, CORS, etc.)
        throw new NetworkError(
          'Failed to connect to SharePoint. Please check your connection.',
          error instanceof Error ? error : undefined
        );
      }

      if (!response.ok) {
        throw createErrorFromResponse(response);
      }

      try {
        const json = await response.json();
        return json.d;
      } catch (error) {
        logError('Failed to parse SharePoint API response', error);
        throw new Error('Invalid response from SharePoint API');
      }
    },
    {
      maxAttempts: 3,
      initialDelayMs: 1000,
      onRetry: (attempt, error) => {
        logDebug(`Retrying SharePoint API request (attempt ${attempt})`, {
          url,
          error: error.message,
        });
      },
    }
  );

  // Store in cache if caching is enabled
  if (cacheOptions) {
    const cacheKey = generateCacheKey(url);
    apiCache.set(cacheKey, data, cacheOptions);
    logDebug('API Cache Stored', { url });
  }

  return data;
}

/**
 * Gets metadata about a specific list or library
 * Cached for 10 minutes since list metadata rarely changes
 */
export async function getListMetadata(webUrl: string, listId: string): Promise<ListMetadata> {
  const url = `${webUrl}/_api/web/lists(guid'${listId}')`;
  const data: any = await fetchSharePointAPI(url, {
    ttl: 10 * 60 * 1000, // 10 minutes
    useSessionStorage: true,
  });

  return {
    id: data.Id,
    title: data.Title,
    internalName: data.EntityTypeName,
    baseTemplate: data.BaseTemplate,
    baseType: data.BaseType,
    itemCount: data.ItemCount,
    description: data.Description || '',
    created: data.Created,
    lastItemModifiedDate: data.LastItemModifiedDate,
    contentTypesEnabled: data.ContentTypesEnabled || false,
    enableVersioning: data.EnableVersioning || false,
  };
}

/**
 * Gets all fields (columns) for a specific list
 * Cached for 10 minutes since field definitions are relatively stable
 */
export async function getListFields(webUrl: string, listId: string): Promise<FieldMetadata[]> {
  const url = `${webUrl}/_api/web/lists(guid'${listId}')/fields?$filter=Hidden eq false`;
  const data: any = await fetchSharePointAPI(url, {
    ttl: 10 * 60 * 1000, // 10 minutes
    useSessionStorage: true,
  });

  return data.results.map((field: any) => ({
    id: field.Id,
    title: field.Title,
    internalName: field.InternalName,
    typeAsString: field.TypeAsString,
    required: field.Required,
    hidden: field.Hidden,
    readOnlyField: field.ReadOnlyField,
    description: field.Description || '',
    group: field.Group || '',
  }));
}

/**
 * Gets all content types for a specific list
 * Cached for 10 minutes since content types are relatively stable
 */
export async function getListContentTypes(
  webUrl: string,
  listId: string
): Promise<ContentTypeMetadata[]> {
  const url = `${webUrl}/_api/web/lists(guid'${listId}')/contenttypes`;
  const data: any = await fetchSharePointAPI(url, {
    ttl: 10 * 60 * 1000, // 10 minutes
    useSessionStorage: true,
  });

  return data.results.map((ct: any) => ({
    id: ct.Id.StringValue,
    name: ct.Name,
    description: ct.Description || '',
    group: ct.Group || '',
  }));
}

/**
 * Gets metadata about the current site collection
 * Cached for 30 minutes since site metadata rarely changes
 */
export async function getSiteMetadata(webUrl: string): Promise<SiteMetadata> {
  const url = `${webUrl}/_api/site`;
  const data: any = await fetchSharePointAPI(url, {
    ttl: 30 * 60 * 1000, // 30 minutes
    useSessionStorage: true,
  });

  return {
    id: data.Id,
    url: data.Url,
    title: '', // Site title is not available directly from site endpoint
    description: '',
    webTemplate: '',
    created: '',
    language: 0,
  };
}

/**
 * Gets metadata about the current web (site or subsite)
 * Cached for 30 minutes since web metadata rarely changes
 */
export async function getWebMetadata(webUrl: string): Promise<WebMetadata> {
  const url = `${webUrl}/_api/web`;
  const data: any = await fetchSharePointAPI(url, {
    ttl: 30 * 60 * 1000, // 30 minutes
    useSessionStorage: true,
  });

  return {
    id: data.Id,
    title: data.Title,
    description: data.Description || '',
    serverRelativeUrl: data.ServerRelativeUrl,
    url: data.Url,
    created: data.Created,
    language: data.Language,
    webTemplate: data.WebTemplate,
  };
}

/**
 * Gets comprehensive metadata for the current context
 */
export async function getContextMetadata(webUrl: string, listId?: string) {
  const results: any = {
    web: null,
    site: null,
    list: null,
    fields: null,
    contentTypes: null,
    error: null,
  };

  try {
    // Always fetch web and site metadata
    results.web = await getWebMetadata(webUrl);
    results.site = await getSiteMetadata(webUrl);

    // If we have a list ID, fetch list-specific metadata
    if (listId) {
      try {
        results.list = await getListMetadata(webUrl, listId);
        results.fields = await getListFields(webUrl, listId);
        results.contentTypes = await getListContentTypes(webUrl, listId);
      } catch (error) {
        logError('Error fetching list metadata', error);
        // Continue even if list metadata fails
      }
    }
  } catch (error) {
    logError('Error fetching SharePoint metadata', error);
    results.error = error instanceof Error ? error.message : 'Unknown error occurred';
  }

  return results;
}

/**
 * Template type mapping for common SharePoint list templates
 */
export function getTemplateTypeName(baseTemplate: number): string {
  const templates: { [key: number]: string } = {
    100: 'Custom List',
    101: 'Document Library',
    102: 'Survey',
    103: 'Links',
    104: 'Announcements',
    105: 'Contacts',
    106: 'Events',
    107: 'Tasks',
    108: 'Discussion Board',
    109: 'Picture Library',
    110: 'Data Sources',
    111: 'Site Template Gallery',
    112: 'User Information List',
    113: 'Web Part Gallery',
    114: 'List Template Gallery',
    115: 'XML Form Library',
    116: 'Master Page Gallery',
    117: 'No-Code Workflows',
    118: 'Custom Workflow Process',
    119: 'Wiki Page Library',
    120: 'Custom Grid',
    130: 'Data Connection Library',
    140: 'Workflow History',
    150: 'Gantt Tasks',
    200: 'Meeting Series',
    201: 'Meeting Agenda',
    202: 'Meeting Attendees',
    204: 'Meeting Decisions',
    207: 'Meeting Objectives',
    210: 'Meeting Text Box',
    211: 'Meeting Things To Bring',
    212: 'Meeting Workspace Pages',
    301: 'Blog Posts',
    302: 'Blog Comments',
    303: 'Blog Categories',
    544: 'Issue Tracking',
    851: 'Page Library',
  };

  return templates[baseTemplate] || `Template ${baseTemplate}`;
}

/**
 * Field type display names
 */
export function getFieldTypeDisplayName(typeAsString: string): string {
  const types: { [key: string]: string } = {
    Text: 'Single line of text',
    Note: 'Multiple lines of text',
    Number: 'Number',
    Currency: 'Currency',
    DateTime: 'Date and Time',
    Boolean: 'Yes/No',
    Choice: 'Choice',
    Lookup: 'Lookup',
    User: 'Person or Group',
    URL: 'Hyperlink',
    Calculated: 'Calculated',
    Attachments: 'Attachments',
    Guid: 'GUID',
    Integer: 'Integer',
    Counter: 'Counter',
    MultiChoice: 'Multiple Choice',
    TaxonomyFieldType: 'Managed Metadata',
    TaxonomyFieldTypeMulti: 'Managed Metadata (Multiple)',
    File: 'File',
    ContentTypeId: 'Content Type ID',
  };

  return types[typeAsString] || typeAsString;
}

/**
 * Generates the URL for editing a specific column
 * Works for both classic and modern SharePoint (redirects to modern if available)
 */
export function getColumnEditUrl(webUrl: string, listId: string, fieldName: string): string {
  // Classic URL format - modern SharePoint will auto-redirect if appropriate
  return `${webUrl}/_layouts/15/FldEdit.aspx?List={${listId}}&Field=${encodeURIComponent(fieldName)}`;
}

/**
 * Generates the URL for managing content types in a list
 * Works for both classic and modern SharePoint
 */
export function getContentTypesManagementUrl(webUrl: string, listId: string): string {
  // Classic URL format for content type settings
  return `${webUrl}/_layouts/15/ManageContentType.aspx?List={${listId}}`;
}

/**
 * Generates the URL for list settings page
 * Works for both classic and modern SharePoint
 */
export function getListSettingsUrl(webUrl: string, listId: string): string {
  return `${webUrl}/_layouts/15/listedit.aspx?List={${listId}}`;
}

/**
 * Site list/library info for navigation
 */
export interface SiteListInfo {
  id: string;
  title: string;
  url: string;
  baseTemplate: number;
  baseType: number;
  isLibrary: boolean;
  hidden: boolean;
  itemCount: number;
}

/**
 * Gets all lists and libraries from the current site
 * Filters out hidden system lists
 * Cached for 5 minutes as lists can be added/removed more frequently
 */
export async function getSiteListsAndLibraries(webUrl: string): Promise<SiteListInfo[]> {
  const url = `${webUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,BaseTemplate,BaseType,ItemCount,Hidden,RootFolder/ServerRelativeUrl&$expand=RootFolder`;

  try {
    const data: any = await fetchSharePointAPI(url, {
      ttl: 5 * 60 * 1000, // 5 minutes
      useSessionStorage: true,
    });

    // Extract origin from webUrl to avoid path duplication
    // ServerRelativeUrl already contains the full path from root (e.g., /sites/sitename/Lists/ListName)
    const origin = new URL(webUrl).origin;

    return data.results.map((list: any) => ({
      id: list.Id,
      title: list.Title,
      url: `${origin}${list.RootFolder.ServerRelativeUrl}`,
      baseTemplate: list.BaseTemplate,
      baseType: list.BaseType,
      isLibrary: list.BaseType === 1, // BaseType 1 = Document Library
      hidden: list.Hidden,
      itemCount: list.ItemCount,
    }));
  } catch (error) {
    logError('Error fetching site lists and libraries', error);
    throw error;
  }
}
