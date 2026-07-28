# SharePoint Context Detection Module

This module provides comprehensive SharePoint context detection capabilities for browser extensions. It can parse SharePoint URLs, detect page types, identify Modern vs Classic UI, and extract list/library information.

## Features

- **URL Parsing**: Extract site collection, web, list IDs, and other components from SharePoint URLs
- **Version Detection**: Identify Modern vs Classic SharePoint experiences
- **Page Context**: Determine the type of page (list, library, site settings, etc.)
- **List Context**: Extract list/library IDs, item IDs, and view IDs
- **Tenant Identification**: Extract tenant information from SharePoint Online URLs

## Usage

### Basic Context Detection

```typescript
import { detectSharePointContext } from './context';

// Detect context from a URL
const result = detectSharePointContext('https://contoso.sharepoint.com/sites/marketing');

if (result.success && result.context) {
  console.log('Site URL:', result.context.urlComponents.webUrl);
  console.log('Version:', result.context.version);
  console.log('Page Type:', result.context.pageContext);
}
```

### In a Browser Extension Content Script

```typescript
import { detectCurrentContext } from './context';

// Detect context from current page
const result = detectCurrentContext(document);

if (result.success && result.context) {
  // Use context information
  const { webUrl, siteCollectionUrl } = result.context.urlComponents;
  const { listId, itemId } = result.context.listContext || {};

  console.log('Current Web:', webUrl);
  console.log('Site Collection:', siteCollectionUrl);

  if (listId) {
    console.log('Current List:', listId);
  }
}
```

### URL Parsing Only

```typescript
import { parseSharePointUrl, extractListId, extractItemId } from './context';

const url = 'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/DispForm.aspx?ID=42';
const components = parseSharePointUrl(url);

if (components) {
  console.log('Web URL:', components.webUrl);
  console.log('List ID:', extractListId(components));
  console.log('Item ID:', extractItemId(components));
}
```

### Check if URL is SharePoint

```typescript
import { isSharePointUrl } from './context';

if (isSharePointUrl(window.location.href)) {
  // This is a SharePoint page
  // Initialize extension features
}
```

## API Reference

### Main Functions

#### `detectSharePointContext(url: string, document?: Document): ContextDetectionResult`

Detects complete SharePoint context from a URL.

**Parameters:**
- `url`: The SharePoint URL to analyze
- `document` (optional): DOM document for enhanced detection

**Returns:** `ContextDetectionResult` with success status and context information

#### `detectCurrentContext(document?: Document): ContextDetectionResult`

Convenience function to detect context from `window.location`.

**Parameters:**
- `document` (optional): DOM document for enhanced detection

**Returns:** `ContextDetectionResult` with success status and context information

#### `parseSharePointUrl(url: string): SharePointUrlComponents | null`

Parses a SharePoint URL and extracts all components.

**Parameters:**
- `url`: The SharePoint URL to parse

**Returns:** URL components or null if not a valid SharePoint URL

#### `isSharePointUrl(url: string): boolean`

Checks if a URL is a SharePoint URL.

**Parameters:**
- `url`: The URL to check

**Returns:** `true` if SharePoint URL, `false` otherwise

### Data Types

#### `SharePointContext`

Complete context information:
```typescript
{
  isSharePoint: boolean;
  version: SharePointVersion;        // Modern, Classic, or Unknown
  pageContext: PageContext;          // SiteHome, List, DocumentLibrary, etc.
  urlComponents: SharePointUrlComponents;
  listContext: ListContext | null;
  detectedAt: Date;
  tenantId: string | null;
}
```

#### `SharePointUrlComponents`

Parsed URL structure:
```typescript
{
  fullUrl: string;
  protocol: string;
  domain: string;
  siteCollectionUrl: string;
  webUrl: string;
  relativePath: string;
  queryParams: Map<string, string>;
}
```

#### `ListContext`

List/Library information:
```typescript
{
  listId: string | null;
  listName: string | null;
  listTitle: string | null;
  itemId: string | null;
  viewId: string | null;
}
```

## Supported URL Patterns

The module recognizes various SharePoint URL patterns:

- **Root Sites**: `https://tenant.sharepoint.com/`
- **Site Collections**: `https://tenant.sharepoint.com/sites/sitename`
- **Team Sites**: `https://tenant.sharepoint.com/teams/teamname`
- **Subsites**: `https://tenant.sharepoint.com/sites/site/subsite`
- **Modern Pages**: `https://tenant.sharepoint.com/sites/site/SitePages/Home.aspx`
- **Document Libraries**: `https://tenant.sharepoint.com/sites/site/Shared Documents`
- **Lists**: `https://tenant.sharepoint.com/sites/site/Lists/ListName`
- **List Forms**: `https://tenant.sharepoint.com/sites/site/Lists/List/DispForm.aspx?ID=1`
- **Settings Pages**: `https://tenant.sharepoint.com/sites/site/_layouts/15/settings.aspx`

## Version Detection

The module detects SharePoint versions using:

1. **URL Patterns**: Identifies Modern vs Classic based on URL structure
2. **DOM Elements** (when document provided): Checks for Modern UI components (suite nav, command bar) vs Classic UI (ribbon)

### Modern Indicators
- `/SitePages/` paths
- Modern command bar elements
- Suite navigation

### Classic Indicators
- Classic forms (`AllItems.aspx`, `DispForm.aspx`, etc.)
- Ribbon UI elements
- Classic layouts

## Page Context Types

The module can identify these page types:

- `SiteHome` - Site home page
- `DocumentLibrary` - Document library view
- `List` - Custom list view
- `ListItem` - List item display/edit form
- `SiteSettings` - Site settings page
- `SiteContents` - Site contents page
- `ModernPage` - Modern SharePoint page
- `WikiPage` - Wiki page
- `WebPartPage` - Classic web part page
- `Unknown` - Unable to determine

## Testing

Run the test suite:

```bash
npm test
```

Or run the test file directly with ts-node:

```bash
npx ts-node tests/context-detector.test.ts
```

## Future Enhancements

- Support for on-premises SharePoint detection
- Enhanced subsite detection with depth tracking
- Content type detection
- Hub site detection
- Associated Microsoft 365 group detection
- Multi-geo tenant support
