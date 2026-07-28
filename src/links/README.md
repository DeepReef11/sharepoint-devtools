# SharePoint Links Registry

This module provides a comprehensive registry of SharePoint navigation links with template-based URL resolution and context-aware filtering.

## Features

- **70+ SharePoint Links** organized into 8 categories
- **Template System** with placeholder support ({webUrl}, {siteUrl}, {listId}, {listUrl})
- **Context-Aware Filtering** based on user permissions and current page context
- **Priority Links** for most commonly used navigation items
- **Search Functionality** across titles, descriptions, and keywords

## Categories

1. **Site Admin** - Settings, permissions, features, regional settings
2. **Content** - Lists, libraries, site contents, recycle bin, columns, content types
3. **Navigation** - Menu structure, quick launch, hub sites
4. **Design** - Themes, branding, master pages, logos
5. **Development** - App catalog, site scripts, SPFx workbench, API access
6. **Modern Admin** - SharePoint admin center, tenant settings
7. **List/Library Settings** - Context-aware settings for current list/library
8. **Advanced** - Search, term store, workflows, information management

## Usage

### Basic Usage

```typescript
import { linkManager, PlaceholderValues } from './links';

// Get all links
const allLinks = linkManager.getAllLinks();

// Get links by category
const adminLinks = linkManager.getLinksByCategory(LinkCategory.SiteAdmin);

// Get priority links (most commonly used)
const priorityLinks = linkManager.getPriorityLinks();
```

### Context-Aware Filtering

```typescript
// Define current context
const values: PlaceholderValues = {
  webUrl: 'https://contoso.sharepoint.com/sites/mysite',
  siteUrl: 'https://contoso.sharepoint.com',
  listId: '{12345678-1234-1234-1234-123456789012}',
  listUrl: 'https://contoso.sharepoint.com/sites/mysite/Lists/MyList',
};

// Get applicable links based on context
const applicableLinks = linkManager.getApplicableLinks(
  values,
  true,  // isListContext
  false, // isSiteAdmin
  false, // isTenantAdmin
  true   // isModern
);
```

### Resolving Link URLs

```typescript
const link = linkManager.getLinkById('site-settings');
const resolvedUrl = linkManager.resolveLink(link, values);
// Result: 'https://contoso.sharepoint.com/sites/mysite/_layouts/15/settings.aspx'
```

### Searching Links

```typescript
// Search by title, description, or keywords
const searchResults = linkManager.search('permissions');
```

### Grouping Links

```typescript
const groupedLinks = linkManager.groupByCategory(applicableLinks);
// Result: { 'Site Admin': [...], 'Content': [...], ... }
```

## Priority Links

The following high-priority links are included:

1. **Site Contents** (priority: 1) - `/_layouts/15/viewlsts.aspx`
2. **List/Library Settings** (priority: 2) - `{listUrl}/_layouts/15/listedit.aspx?List={listId}` (context-aware)
3. **Site Settings** (priority: 3) - `/_layouts/15/settings.aspx`
4. **Site Permissions** (priority: 4) - `/_layouts/15/user.aspx`
5. **Site Columns** (priority: 5) - `/_layouts/15/mngfield.aspx`
6. **Site Content Types** (priority: 6) - `/_layouts/15/mngctype.aspx`
7. **Recycle Bin** (priority: 7) - `/_layouts/15/RecycleBin.aspx`
8. **Site Collection Recycle Bin** (priority: 8) - `/_layouts/15/AdminRecycleBin.aspx`

## Placeholder Types

- `{webUrl}` - Current web URL (e.g., `https://tenant.sharepoint.com/sites/sitename`)
- `{siteUrl}` - Site collection URL (e.g., `https://tenant.sharepoint.com`)
- `{listId}` - Current list/library GUID
- `{listUrl}` - Current list/library root URL

## Link Structure

Each link has the following structure:

```typescript
interface SharePointLink {
  id: string;                    // Unique identifier
  title: string;                 // Display name
  urlTemplate: string;           // URL with placeholders
  category: LinkCategory;        // Grouping category
  description: string;           // What the link does
  context?: LinkContext;         // Applicability requirements
  keywords?: string[];           // Search terms
  priority?: number;             // Display priority (lower = higher)
}
```

## Context Requirements

Links can have context requirements that determine when they're applicable:

```typescript
interface LinkContext {
  requiresList?: boolean;        // Must be in list/library context
  requiresSiteAdmin?: boolean;   // Must have site admin permissions
  requiresTenantAdmin?: boolean; // Must have tenant admin permissions
  modernOnly?: boolean;          // Only for modern SharePoint
  classicOnly?: boolean;         // Only for classic SharePoint
}
```

## Examples

### Example 1: Get Links for Current Page

```typescript
import { linkManager, PlaceholderValues } from './links';

// Extract context from current page
const values: PlaceholderValues = {
  webUrl: window.location.origin + window.location.pathname.split('/_layouts')[0],
  siteUrl: window.location.origin,
};

// Check if we're in a list context
const isListContext = window.location.href.includes('/Lists/') ||
                      window.location.href.includes('/Forms/');

// Get applicable links
const links = linkManager.getApplicableLinks(
  values,
  isListContext,
  false, // Assume not site admin for now
  false, // Assume not tenant admin
  true   // Assume modern SharePoint
);

// Sort by priority
const sortedLinks = linkManager.sortByPriority(links);
```

### Example 2: Search and Resolve

```typescript
// Search for recycle bin links
const searchResults = linkManager.search('recycle bin');

// Resolve URLs for results
const resolvedLinks = searchResults.map(link => ({
  ...link,
  url: linkManager.resolveLink(link, values)
})).filter(link => link.url !== null);
```

## Integration with Other Modules

This module is designed to work with:

- **Context Detection Module** - Provides placeholder values
- **Link Resolution Engine** - Uses the template engine
- **Fuzzy Search** - Will integrate with the search functionality
- **Modal UI** - Will display the links

## Future Enhancements

- Integration with Fuse.js for advanced fuzzy search
- Custom user-defined links
- Recent/favorite link tracking
- Dynamic link validation
- Performance optimization with lazy loading
