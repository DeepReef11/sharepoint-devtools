import { SharePointLink, LinkCategory } from './types';

/**
 * Comprehensive registry of SharePoint navigation links
 * Organized by category with priority links marked
 */
export const SHAREPOINT_LINKS: SharePointLink[] = [
  // ============================================================================
  // SITE ADMIN CATEGORY
  // ============================================================================
  {
    id: 'site-settings',
    title: 'Site Settings',
    urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'Main site settings page - gateway to all configuration options',
    keywords: ['settings', 'configuration', 'admin', 'site'],
    priority: 9,
  },
  {
    id: 'site-permissions',
    title: 'Site Permissions',
    urlTemplate: '{webUrl}/_layouts/15/user.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'Manage site access and permissions',
    keywords: ['permissions', 'access', 'security', 'users', 'groups'],
    priority: 4,
  },
  {
    id: 'people-and-groups',
    title: 'People and Groups',
    urlTemplate: '{webUrl}/_layouts/15/people.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'Manage site users, groups, and membership',
    keywords: ['people', 'groups', 'users', 'members', 'membership', 'permissions'],
    priority: 5,
  },
  {
    id: 'regional-settings',
    title: 'Regional Settings',
    urlTemplate: '{webUrl}/_layouts/15/regionalsetng.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'Configure locale, time zone, and calendar settings',
    keywords: ['regional', 'locale', 'timezone', 'calendar'],
  },
  {
    id: 'site-features',
    title: 'Site Features',
    urlTemplate: '{webUrl}/_layouts/15/ManageFeatures.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'Activate or deactivate site features',
    keywords: ['features', 'activate', 'deactivate'],
  },
  {
    id: 'site-collection-features',
    title: 'Site Collection Features',
    urlTemplate: '{siteUrl}/_layouts/15/ManageFeatures.aspx?Scope=Site',
    category: LinkCategory.SiteAdmin,
    description: 'Manage site collection-level features',
    keywords: ['features', 'site collection', 'activate'],
    context: {
      requiresSiteAdmin: true,
    },
  },
  {
    id: 'site-usage',
    title: 'Site Usage Analytics',
    urlTemplate: '{webUrl}/_layouts/15/usagedetails.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'View site usage statistics and analytics',
    keywords: ['usage', 'analytics', 'statistics', 'reports'],
  },
  {
    id: 'storage-metrics',
    title: 'Storage Metrics',
    urlTemplate: '{siteUrl}/_layouts/15/storman.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'View storage usage by list and library',
    keywords: ['storage', 'space', 'quota', 'metrics'],
    context: {
      requiresSiteAdmin: true,
    },
  },

  // ============================================================================
  // CONTENT CATEGORY
  // ============================================================================
  {
    id: 'site-contents',
    title: 'Site Contents',
    urlTemplate: '{webUrl}/_layouts/15/viewlsts.aspx',
    category: LinkCategory.Content,
    description: 'View all lists, libraries, and apps in the site',
    keywords: ['contents', 'lists', 'libraries', 'apps', 'all'],
    priority: 10,
  },
  {
    id: 'recycle-bin',
    title: 'Recycle Bin',
    urlTemplate: '{webUrl}/_layouts/15/RecycleBin.aspx',
    category: LinkCategory.Content,
    description: 'Recover deleted items from the site recycle bin',
    keywords: ['recycle', 'bin', 'deleted', 'restore', 'recover'],
    priority: 7,
  },
  {
    id: 'site-collection-recycle-bin',
    title: 'Site Collection Recycle Bin',
    urlTemplate: '{siteUrl}/_layouts/15/AdminRecycleBin.aspx',
    category: LinkCategory.Content,
    description: 'Recover deleted items from site collection (second-stage) recycle bin (admin)',
    keywords: [
      'recycle',
      'bin',
      'site collection',
      'admin',
      'restore',
      'second stage',
      'second-stage',
    ],
    priority: 8,
  },
  {
    id: 'site-columns',
    title: 'Site Columns',
    urlTemplate: '{webUrl}/_layouts/15/mngfield.aspx',
    category: LinkCategory.Content,
    description: 'Manage reusable columns across the site',
    keywords: ['columns', 'fields', 'metadata', 'site columns'],
    priority: 5,
  },
  {
    id: 'site-content-types',
    title: 'Site Content Types',
    urlTemplate: '{webUrl}/_layouts/15/mngctype.aspx',
    category: LinkCategory.Content,
    description: 'Define and manage content types',
    keywords: ['content types', 'templates', 'schema'],
    priority: 6,
  },
  {
    id: 'create-list',
    title: 'Create List',
    urlTemplate: '{webUrl}/_layouts/15/new.aspx',
    category: LinkCategory.Content,
    description: 'Create a new list or library',
    keywords: ['create', 'new', 'list', 'library'],
  },
  {
    id: 'pages-library',
    title: 'Site Pages Library',
    urlTemplate: '{webUrl}/SitePages/Forms/AllPages.aspx',
    category: LinkCategory.Content,
    description: 'Manage site pages',
    keywords: ['pages', 'library', 'site pages'],
  },

  // ============================================================================
  // NAVIGATION CATEGORY
  // ============================================================================
  {
    id: 'navigation-settings',
    title: 'Navigation Settings',
    urlTemplate: '{webUrl}/_layouts/15/AreaNavigationSettings.aspx',
    category: LinkCategory.Navigation,
    description: 'Configure site navigation and menu structure',
    keywords: ['navigation', 'menu', 'links', 'structure'],
  },
  {
    id: 'top-navigation',
    title: 'Top Navigation',
    urlTemplate: '{webUrl}/_layouts/15/topnav.aspx',
    category: LinkCategory.Navigation,
    description: 'Edit top navigation menu',
    keywords: ['top', 'navigation', 'menu', 'header'],
  },
  {
    id: 'quick-launch',
    title: 'Quick Launch',
    urlTemplate: '{webUrl}/_layouts/15/quiklnch.aspx',
    category: LinkCategory.Navigation,
    description: 'Edit quick launch (left navigation)',
    keywords: ['quick launch', 'left', 'navigation', 'sidebar'],
  },
  {
    id: 'tree-view',
    title: 'Tree View Settings',
    urlTemplate: '{webUrl}/_layouts/15/navoptions.aspx',
    category: LinkCategory.Navigation,
    description: 'Enable/disable tree view navigation',
    keywords: ['tree', 'view', 'navigation'],
  },
  {
    id: 'hub-site-settings',
    title: 'Hub Site Settings',
    urlTemplate: '{webUrl}/_layouts/15/hubsitedata.aspx',
    category: LinkCategory.Navigation,
    description: 'Configure hub site associations',
    keywords: ['hub', 'site', 'association'],
    context: {
      modernOnly: true,
    },
  },

  // ============================================================================
  // DESIGN CATEGORY
  // ============================================================================
  {
    id: 'change-look',
    title: 'Change the Look',
    urlTemplate: '{webUrl}/_layouts/15/designgallery.aspx',
    category: LinkCategory.Design,
    description: 'Change site theme and appearance',
    keywords: ['theme', 'look', 'design', 'appearance', 'branding'],
  },
  {
    id: 'site-theme',
    title: 'Site Theme',
    urlTemplate: '{webUrl}/_layouts/15/themeweb.aspx',
    category: LinkCategory.Design,
    description: 'Apply or customize site theme',
    keywords: ['theme', 'colors', 'branding'],
  },
  {
    id: 'master-page-gallery',
    title: 'Master Page Gallery',
    urlTemplate: '{webUrl}/_catalogs/masterpage/Forms/AllItems.aspx',
    category: LinkCategory.Design,
    description: 'Manage master pages and page layouts',
    keywords: ['master', 'page', 'layout', 'gallery'],
    context: {
      classicOnly: true,
    },
  },
  {
    id: 'site-logo',
    title: 'Title, Description, and Logo',
    urlTemplate: '{webUrl}/_layouts/15/prjsetng.aspx',
    category: LinkCategory.Design,
    description: 'Change site title, description, and logo',
    keywords: ['title', 'description', 'logo', 'icon'],
  },
  {
    id: 'welcome-page',
    title: 'Welcome Page',
    urlTemplate: '{webUrl}/_layouts/15/AreaWelcomePage.aspx',
    category: LinkCategory.Design,
    description: 'Set the site home page',
    keywords: ['welcome', 'home', 'page', 'default'],
  },

  // ============================================================================
  // DEVELOPMENT CATEGORY
  // ============================================================================
  {
    id: 'app-catalog',
    title: 'App Catalog (Tenant)',
    urlTemplate: '{siteUrl}/_layouts/15/tenantAppCatalog.aspx',
    category: LinkCategory.Development,
    description: 'Tenant-level app catalog - manage SharePoint apps and solutions',
    keywords: ['app', 'catalog', 'solutions', 'spfx', 'tenant', 'global'],
    priority: 6,
  },
  {
    id: 'site-app-catalog',
    title: 'App Catalog (Site)',
    urlTemplate: '{webUrl}/_layouts/15/AppCatalog.aspx',
    category: LinkCategory.Development,
    description: 'Site collection app catalog - manage site-specific apps',
    keywords: ['app', 'catalog', 'site', 'collection', 'local'],
    priority: 6,
  },
  {
    id: 'site-scripts',
    title: 'Site Scripts',
    urlTemplate: '{siteUrl}/_layouts/15/SiteScripts.aspx',
    category: LinkCategory.Development,
    description: 'Manage site scripts for provisioning',
    keywords: ['site', 'scripts', 'provisioning', 'json'],
    context: {
      requiresTenantAdmin: true,
      modernOnly: true,
    },
  },
  {
    id: 'site-designs',
    title: 'Site Designs',
    urlTemplate: '{siteUrl}/_layouts/15/SiteDesigns.aspx',
    category: LinkCategory.Development,
    description: 'Manage site designs and templates',
    keywords: ['site', 'designs', 'templates', 'provisioning'],
    context: {
      requiresTenantAdmin: true,
      modernOnly: true,
    },
  },
  {
    id: 'api-access',
    title: 'API Access',
    urlTemplate: '{siteUrl}/_layouts/15/appinv.aspx',
    category: LinkCategory.Development,
    description: 'Manage API permissions and app principals',
    keywords: ['api', 'access', 'permissions', 'app'],
  },
  {
    id: 'spfx-workbench',
    title: 'SPFx Workbench',
    urlTemplate: '{webUrl}/_layouts/15/workbench.aspx',
    category: LinkCategory.Development,
    description: 'SharePoint Framework development workbench (hosted)',
    keywords: ['spfx', 'workbench', 'development', 'framework', 'hosted'],
    priority: 5,
    context: {
      modernOnly: true,
    },
  },
  {
    id: 'spfx-local-workbench',
    title: 'SPFx Local Workbench',
    urlTemplate: 'https://localhost:4321/temp/workbench.html',
    category: LinkCategory.Development,
    description: 'SharePoint Framework local development workbench (requires gulp serve)',
    keywords: ['spfx', 'workbench', 'development', 'framework', 'local', 'localhost'],
    priority: 5,
  },
  {
    id: 'solutions-gallery',
    title: 'Solutions Gallery',
    urlTemplate: '{siteUrl}/_catalogs/solutions/Forms/AllItems.aspx',
    category: LinkCategory.Development,
    description: 'Manage sandboxed solutions',
    keywords: ['solutions', 'gallery', 'wsp'],
    context: {
      requiresSiteAdmin: true,
    },
  },

  // ============================================================================
  // MODERN ADMIN CATEGORY
  // ============================================================================
  {
    id: 'sharepoint-admin-center',
    title: 'SharePoint Admin Center',
    urlTemplate: '{tenantAdminUrl}/_layouts/15/online/AdminHome.aspx',
    category: LinkCategory.ModernAdmin,
    description: 'SharePoint Online admin center - tenant administration',
    keywords: ['admin', 'center', 'sharepoint', 'tenant', 'administration'],
    priority: 7,
  },
  {
    id: 'microsoft-365-admin-sharepoint',
    title: 'SharePoint (Microsoft 365 Admin)',
    urlTemplate: 'https://admin.microsoft.com/sharepoint',
    category: LinkCategory.ModernAdmin,
    description: 'SharePoint settings in Microsoft 365 admin center',
    keywords: ['admin', 'center', 'microsoft 365', 'm365', 'sharepoint', 'modern'],
    priority: 7,
  },
  {
    id: 'active-sites',
    title: 'Active Sites',
    urlTemplate: '{siteUrl}/_layouts/15/online/SitesList.aspx',
    category: LinkCategory.ModernAdmin,
    description: 'View all active sites in the tenant',
    keywords: ['sites', 'active', 'tenant', 'admin'],
    context: {
      requiresTenantAdmin: true,
      modernOnly: true,
    },
  },
  {
    id: 'deleted-sites',
    title: 'Deleted Sites',
    urlTemplate: '{siteUrl}/_layouts/15/online/DeletedSites.aspx',
    category: LinkCategory.ModernAdmin,
    description: 'Restore deleted sites',
    keywords: ['deleted', 'sites', 'restore', 'recycle'],
    context: {
      requiresTenantAdmin: true,
      modernOnly: true,
    },
  },
  {
    id: 'sharing-settings',
    title: 'Sharing Settings',
    urlTemplate: '{siteUrl}/_layouts/15/online/ExternalSharing.aspx',
    category: LinkCategory.ModernAdmin,
    description: 'Configure external sharing settings',
    keywords: ['sharing', 'external', 'settings', 'security'],
    context: {
      requiresTenantAdmin: true,
      modernOnly: true,
    },
  },
  {
    id: 'tenant-settings',
    title: 'Tenant Settings',
    urlTemplate: '{siteUrl}/_layouts/15/online/TenantSettings.aspx',
    category: LinkCategory.ModernAdmin,
    description: 'Manage tenant-wide settings',
    keywords: ['tenant', 'settings', 'admin'],
    context: {
      requiresTenantAdmin: true,
      modernOnly: true,
    },
  },

  // ============================================================================
  // LIST/LIBRARY SETTINGS CATEGORY
  // ============================================================================
  {
    id: 'list-settings',
    title: 'List/Library Settings',
    urlTemplate: '{listUrl}/_layouts/15/listedit.aspx?List={listId}',
    category: LinkCategory.ListLibrarySettings,
    description: 'Main settings page for current list or library',
    keywords: ['list', 'library', 'settings', 'configuration'],
    priority: 15,
    context: {
      requiresList: true,
    },
  },
  {
    id: 'list-permissions',
    title: 'List Permissions',
    urlTemplate: '{listUrl}/_layouts/15/user.aspx?List={listId}',
    category: LinkCategory.ListLibrarySettings,
    description: 'Manage permissions for current list or library',
    keywords: ['permissions', 'access', 'security', 'list'],
    context: {
      requiresList: true,
    },
  },
  {
    id: 'list-columns',
    title: 'List Columns',
    urlTemplate: '{listUrl}/_layouts/15/listedit.aspx?List={listId}&page=3',
    category: LinkCategory.ListLibrarySettings,
    description: 'Manage columns in current list or library',
    keywords: ['columns', 'fields', 'list', 'library'],
    context: {
      requiresList: true,
    },
  },
  {
    id: 'list-views',
    title: 'List Views',
    urlTemplate: '{listUrl}/_layouts/15/viewtype.aspx?List={listId}',
    category: LinkCategory.ListLibrarySettings,
    description: 'Create and manage views',
    keywords: ['views', 'list', 'library'],
    context: {
      requiresList: true,
    },
  },
  {
    id: 'list-versioning',
    title: 'Versioning Settings',
    urlTemplate: '{listUrl}/_layouts/15/listedit.aspx?List={listId}&page=5',
    category: LinkCategory.ListLibrarySettings,
    description: 'Configure version history and approval',
    keywords: ['versioning', 'history', 'approval', 'list'],
    context: {
      requiresList: true,
    },
  },
  {
    id: 'list-validation',
    title: 'Validation Settings',
    urlTemplate: '{listUrl}/_layouts/15/listedit.aspx?List={listId}&page=8',
    category: LinkCategory.ListLibrarySettings,
    description: 'Set up list validation rules',
    keywords: ['validation', 'rules', 'list'],
    context: {
      requiresList: true,
    },
  },
  {
    id: 'list-workflow',
    title: 'Workflow Settings',
    urlTemplate: '{listUrl}/_layouts/15/AddAnAppForListDialog.aspx?List={listId}',
    category: LinkCategory.ListLibrarySettings,
    description: 'Add and manage workflows',
    keywords: ['workflow', 'automation', 'list'],
    context: {
      requiresList: true,
    },
  },
  {
    id: 'list-content-types',
    title: 'Advanced Settings',
    urlTemplate: '{listUrl}/_layouts/15/listedit.aspx?List={listId}&page=1',
    category: LinkCategory.ListLibrarySettings,
    description: 'Advanced list/library settings including content types',
    keywords: ['advanced', 'settings', 'content types', 'list'],
    context: {
      requiresList: true,
    },
  },

  // ============================================================================
  // ADVANCED CATEGORY
  // ============================================================================
  {
    id: 'search-settings',
    title: 'Search Settings',
    urlTemplate: '{webUrl}/_layouts/15/enhancedSearch.aspx',
    category: LinkCategory.Advanced,
    description: 'Configure site search settings',
    keywords: ['search', 'settings', 'indexing'],
  },
  {
    id: 'search-schema',
    title: 'Search Schema',
    urlTemplate: '{webUrl}/_layouts/15/listmanagedproperties.aspx?level=sitecol',
    category: LinkCategory.Advanced,
    description: 'Manage search schema and managed properties (site collection level)',
    keywords: ['search', 'schema', 'managed', 'properties', 'crawled'],
    priority: 6,
  },
  {
    id: 'term-store',
    title: 'Term Store Management',
    urlTemplate: '{siteUrl}/_layouts/15/termstoremanager.aspx',
    category: LinkCategory.Advanced,
    description: 'Manage taxonomy, term sets, and metadata',
    keywords: ['term', 'store', 'taxonomy', 'metadata', 'managed metadata'],
    priority: 6,
  },
  {
    id: 'information-management',
    title: 'Information Management Policy',
    urlTemplate: '{webUrl}/_layouts/15/Policylist.aspx',
    category: LinkCategory.Advanced,
    description: 'Configure retention and auditing policies',
    keywords: ['policy', 'retention', 'compliance', 'auditing'],
  },
  {
    id: 'audit-log',
    title: 'Audit Log Reports',
    urlTemplate: '{siteUrl}/_layouts/15/Reporting.aspx',
    category: LinkCategory.Advanced,
    description: 'View audit log reports',
    keywords: ['audit', 'log', 'reports', 'compliance'],
    context: {
      requiresSiteAdmin: true,
    },
  },
  {
    id: 'site-closure',
    title: 'Site Closure and Deletion',
    urlTemplate: '{webUrl}/_layouts/15/sitemanager.aspx',
    category: LinkCategory.Advanced,
    description: 'Configure site closure policy',
    keywords: ['closure', 'deletion', 'policy', 'site'],
    context: {
      requiresSiteAdmin: true,
    },
  },
  {
    id: 'web-parts',
    title: 'Web Part Gallery',
    urlTemplate: '{siteUrl}/_catalogs/wp/Forms/AllItems.aspx',
    category: LinkCategory.Advanced,
    description: 'Manage web part definitions',
    keywords: ['web', 'parts', 'gallery'],
    context: {
      requiresSiteAdmin: true,
    },
  },
  {
    id: 'list-templates',
    title: 'List Template Gallery',
    urlTemplate: '{webUrl}/_catalogs/lt/Forms/AllItems.aspx',
    category: LinkCategory.Advanced,
    description: 'Manage list templates',
    keywords: ['list', 'templates', 'gallery'],
  },
];
