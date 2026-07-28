/**
 * Sample Links Registry
 * Example link templates for testing the Link Resolution Engine
 * This will be expanded in Issue #4: SharePoint Links Registry
 */

import { LinkTemplate, LinkCategory, LinkRegistry } from '../types/link-template';
import { SharePointPageType, SharePointVersion } from '../types/sharepoint-context';
import { PlaceholderType } from '../types/link-template';

/**
 * Sample link templates for testing
 */
export const sampleLinkTemplates: LinkTemplate[] = [
  // Site Admin Links
  {
    id: 'site-settings',
    title: 'Site Settings',
    urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'Access site settings and configuration',
    keywords: ['settings', 'configuration', 'admin'],
    applicability: {
      pageTypes: [], // Available on all page types
      versions: [], // Available on all versions
      requiredContext: [PlaceholderType.WebUrl],
    },
    openInNewTab: false,
  },
  {
    id: 'site-permissions',
    title: 'Site Permissions',
    urlTemplate: '{webUrl}/_layouts/15/user.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'Manage site permissions and users',
    keywords: ['permissions', 'users', 'access', 'security'],
    applicability: {
      pageTypes: [],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl],
    },
    openInNewTab: false,
  },
  {
    id: 'site-features',
    title: 'Site Features',
    urlTemplate: '{webUrl}/_layouts/15/ManageFeatures.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'Activate or deactivate site features',
    keywords: ['features', 'activate', 'deactivate'],
    applicability: {
      pageTypes: [],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl],
    },
    openInNewTab: false,
  },

  // Content Links
  {
    id: 'site-contents',
    title: 'Site Contents',
    urlTemplate: '{webUrl}/_layouts/15/viewlsts.aspx',
    category: LinkCategory.Content,
    description: 'View all lists and libraries',
    keywords: ['contents', 'lists', 'libraries', 'all'],
    applicability: {
      pageTypes: [],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl],
    },
    openInNewTab: false,
  },
  {
    id: 'recycle-bin',
    title: 'Recycle Bin',
    urlTemplate: '{webUrl}/_layouts/15/RecycleBin.aspx',
    category: LinkCategory.Content,
    description: 'View deleted items',
    keywords: ['recycle', 'bin', 'deleted', 'trash'],
    applicability: {
      pageTypes: [],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl],
    },
    openInNewTab: false,
  },

  // List/Library Settings (only applicable on list pages)
  {
    id: 'list-settings',
    title: 'List Settings',
    urlTemplate: '{webUrl}/_layouts/15/listedit.aspx?List={listId}',
    category: LinkCategory.ListSettings,
    description: 'Configure list settings',
    keywords: ['list', 'settings', 'configuration'],
    applicability: {
      pageTypes: [SharePointPageType.List, SharePointPageType.Library],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl, PlaceholderType.ListId],
    },
    openInNewTab: false,
  },
  {
    id: 'list-permissions',
    title: 'List Permissions',
    urlTemplate: '{webUrl}/_layouts/15/user.aspx?List={listId}',
    category: LinkCategory.ListSettings,
    description: 'Manage list permissions',
    keywords: ['list', 'permissions', 'access'],
    applicability: {
      pageTypes: [SharePointPageType.List, SharePointPageType.Library],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl, PlaceholderType.ListId],
    },
    openInNewTab: false,
  },
  {
    id: 'list-columns',
    title: 'List Columns',
    urlTemplate: '{webUrl}/_layouts/15/ListEdit.aspx?List={listId}&view=Columns',
    category: LinkCategory.ListSettings,
    description: 'Manage list columns',
    keywords: ['columns', 'fields', 'metadata'],
    applicability: {
      pageTypes: [SharePointPageType.List, SharePointPageType.Library],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl, PlaceholderType.ListId],
    },
    openInNewTab: false,
  },
  {
    id: 'list-views',
    title: 'List Views',
    urlTemplate: '{webUrl}/_layouts/15/ViewEdit.aspx?List={listId}',
    category: LinkCategory.ListSettings,
    description: 'Manage list views',
    keywords: ['views', 'filters', 'sorting'],
    applicability: {
      pageTypes: [SharePointPageType.List, SharePointPageType.Library],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl, PlaceholderType.ListId],
    },
    openInNewTab: false,
  },

  // Design Links
  {
    id: 'change-look',
    title: 'Change the Look',
    urlTemplate: '{webUrl}/_layouts/15/designpicker.aspx',
    category: LinkCategory.Design,
    description: 'Apply site themes and designs',
    keywords: ['theme', 'design', 'branding', 'look'],
    applicability: {
      pageTypes: [],
      versions: [SharePointVersion.Classic], // Classic only
      requiredContext: [PlaceholderType.WebUrl],
    },
    openInNewTab: false,
  },
  {
    id: 'site-pages-library',
    title: 'Site Pages Library',
    urlTemplate: '{webUrl}/SitePages/Forms/AllPages.aspx',
    category: LinkCategory.Design,
    description: 'Manage site pages',
    keywords: ['pages', 'library', 'content'],
    applicability: {
      pageTypes: [],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl],
    },
    openInNewTab: false,
  },

  // Modern Admin Links
  {
    id: 'sharepoint-admin',
    title: 'SharePoint Admin Center',
    urlTemplate: '{tenantUrl}/_layouts/15/online/AdminHome.aspx',
    category: LinkCategory.ModernAdmin,
    description: 'Open SharePoint Admin Center',
    keywords: ['admin', 'center', 'tenant', 'management'],
    applicability: {
      pageTypes: [],
      versions: [],
      requiredContext: [PlaceholderType.TenantUrl],
    },
    openInNewTab: true,
  },

  // Advanced Links
  {
    id: 'search-settings',
    title: 'Search Settings',
    urlTemplate: '{webUrl}/_layouts/15/enhancedSearch.aspx',
    category: LinkCategory.Advanced,
    description: 'Configure search settings',
    keywords: ['search', 'settings', 'configuration'],
    applicability: {
      pageTypes: [],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl],
    },
    openInNewTab: false,
  },
  {
    id: 'regional-settings',
    title: 'Regional Settings',
    urlTemplate: '{webUrl}/_layouts/15/regionalsetng.aspx',
    category: LinkCategory.Advanced,
    description: 'Configure regional settings (timezone, locale)',
    keywords: ['regional', 'timezone', 'locale', 'language'],
    applicability: {
      pageTypes: [],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl],
    },
    openInNewTab: false,
  },
];

/**
 * Sample link registry
 */
export const sampleLinkRegistry: LinkRegistry = {
  templates: sampleLinkTemplates,
  version: '1.0.0',
};
