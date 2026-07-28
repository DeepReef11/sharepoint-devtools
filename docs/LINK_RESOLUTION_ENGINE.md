# Link Resolution Engine

## Overview

The Link Resolution Engine is the core component responsible for converting link templates with placeholders into fully resolved, context-aware URLs for SharePoint navigation.

## Features

### 1. Placeholder Replacement

Supports the following placeholders:
- `{tenantUrl}` - SharePoint tenant root URL (e.g., `https://contoso.sharepoint.com`)
- `{siteUrl}` - Site collection URL (e.g., `https://contoso.sharepoint.com/sites/mysite`)
- `{webUrl}` - Web/subsite URL (currently same as siteUrl, subsites in future)
- `{listId}` - List or library GUID
- `{listTitle}` - List or library title

**Example:**
```typescript
const template = '{webUrl}/_layouts/15/settings.aspx';
const context = detectSharePointContext(window.location.href);
const resolved = LinkResolver.replacePlaceholders(template, context);
// Result: https://contoso.sharepoint.com/sites/mysite/_layouts/15/settings.aspx
```

### 2. Link Applicability Logic (Context-Aware Filtering)

Links are filtered based on:
- **Page Type**: Site, List, Library, Page
- **SharePoint Version**: Modern vs Classic
- **Required Context**: Which placeholders must be available

**Example:**
```typescript
const listSettingsLink: LinkTemplate = {
  id: 'list-settings',
  title: 'List Settings',
  urlTemplate: '{webUrl}/_layouts/15/listedit.aspx?List={listId}',
  category: LinkCategory.ListSettings,
  applicability: {
    pageTypes: [SharePointPageType.List, SharePointPageType.Library],
    requiredContext: [PlaceholderType.WebUrl, PlaceholderType.ListId],
  },
};

// Only shows when on a list/library page with listId available
const isApplicable = LinkResolver.isLinkApplicable(listSettingsLink, context);
```

### 3. Link Validation

Validates resolved URLs to ensure:
- All placeholders are replaced (no `{placeholder}` remains)
- URL is valid and well-formed
- URL is on the same tenant or a trusted admin URL

**Example:**
```typescript
const url = 'https://contoso.sharepoint.com/sites/mysite/_layouts/15/settings.aspx';
const validation = LinkResolver.validateUrl(url, context);
// { isValid: true, errors: [] }
```

### 4. Final URL Generation

Combines all the above to generate fully resolved, validated links:

**Example:**
```typescript
const resolved = LinkResolver.resolveLink(template, context);
// Returns ResolvedLink with:
// - url: fully resolved URL
// - isValid: validation result
// - validationErrors: any errors found
```

## API Reference

### LinkResolver Class

#### `resolveLink(template: LinkTemplate, context: SharePointContext): ResolvedLink | null`

Resolves a single link template. Returns `null` if not applicable in current context.

#### `resolveLinks(templates: LinkTemplate[], context: SharePointContext): ResolvedLink[]`

Resolves multiple templates at once.

#### `isLinkApplicable(template: LinkTemplate, context: SharePointContext): boolean`

Checks if a link should be shown in the current context.

#### `replacePlaceholders(urlTemplate: string, context: SharePointContext): string`

Replaces all placeholders in a URL template.

#### `validateUrl(url: string, context: SharePointContext): { isValid: boolean; errors: string[] }`

Validates a resolved URL.

#### `filterApplicableTemplates(templates: LinkTemplate[], context: SharePointContext): LinkTemplate[]`

Returns only templates applicable in current context.

#### `groupByCategory(links: ResolvedLink[]): Map<string, ResolvedLink[]>`

Groups resolved links by category.

#### `extractPlaceholders(urlTemplate: string): string[]`

Extracts all placeholder names from a template.

#### `getMissingContext(template: LinkTemplate, context: SharePointContext): string[]`

Returns which placeholders are missing from the current context.

## Context Detection

### SharePointContext Interface

```typescript
interface SharePointContext {
  tenantUrl: string | null;
  siteUrl: string | null;
  webUrl: string | null;
  listId: string | null;
  listTitle: string | null;
  pageType: SharePointPageType;
  version: SharePointVersion;
  isSharePointPage: boolean;
  currentUrl: string;
}
```

### Context Detection Functions

#### `detectSharePointContext(url?: string): SharePointContext`

Detects SharePoint context from a URL (defaults to current page).

#### `getCurrentContext(): SharePointContext`

Gets context for the current page (browser only).

#### `isSharePointUrl(url: string): boolean`

Checks if a URL is a SharePoint URL.

## Usage Examples

### Basic Usage

```typescript
import { LinkResolver, detectSharePointContext, sampleLinkTemplates } from './src/index';

// Detect current context
const context = detectSharePointContext(window.location.href);

// Resolve all applicable links
const resolvedLinks = LinkResolver.resolveLinks(sampleLinkTemplates, context);

// Group by category
const grouped = LinkResolver.groupByCategory(resolvedLinks);

// Display links
grouped.forEach((links, category) => {
  console.log(`${category}:`);
  links.forEach(link => {
    console.log(`  - ${link.title}: ${link.url}`);
  });
});
```

### Creating Custom Link Templates

```typescript
import { LinkTemplate, LinkCategory, PlaceholderType, SharePointPageType } from './src/index';

const customLink: LinkTemplate = {
  id: 'my-custom-link',
  title: 'My Custom Link',
  urlTemplate: '{webUrl}/_layouts/15/mypage.aspx',
  category: LinkCategory.SiteAdmin,
  description: 'My custom SharePoint page',
  keywords: ['custom', 'page'],
  applicability: {
    pageTypes: [SharePointPageType.Site],
    requiredContext: [PlaceholderType.WebUrl],
  },
  openInNewTab: false,
};
```

### Filtering for Specific Context

```typescript
// Get only list-specific links
const listContext = detectSharePointContext(
  'https://contoso.sharepoint.com/sites/mysite/Lists/MyList/AllItems.aspx'
);

const listLinks = LinkResolver.resolveLinks(sampleLinkTemplates, listContext);
console.log(`Found ${listLinks.length} applicable links for this list`);
```

## Testing

Run the test suite:

```typescript
import { runAllTests } from './src/tests/link-resolver.test';

// In browser console:
runAllTests();

// Or run individual tests:
import { testPlaceholderReplacement } from './src/tests/link-resolver.test';
testPlaceholderReplacement();
```

## Architecture

```
┌─────────────────────────────────────────────────────────┐
│                    Link Resolution Engine                │
└─────────────────────────────────────────────────────────┘
                            │
        ┌───────────────────┼───────────────────┐
        │                   │                   │
        ▼                   ▼                   ▼
┌───────────────┐  ┌─────────────────┐  ┌─────────────┐
│   Context     │  │  Link Templates │  │  Validation │
│   Detection   │  │   (Registry)    │  │   Engine    │
└───────────────┘  └─────────────────┘  └─────────────┘
        │                   │                   │
        └───────────────────┼───────────────────┘
                            ▼
                    ┌─────────────────┐
                    │ Resolved Links  │
                    │ (with URLs)     │
                    └─────────────────┘
```

## Future Enhancements

1. **Subsite Detection**: Currently `webUrl` = `siteUrl`. Future: detect subsites.
2. **Permission Checking**: Validate user permissions before showing links.
3. **Dynamic Context**: Listen for navigation changes and update context.
4. **Link Caching**: Cache resolved links for performance.
5. **Custom Validators**: Allow custom validation rules per link.

## Dependencies

This implementation satisfies Issue #5 and partially implements Issue #3 (SharePoint Context Detection).

### Related Issues

- SharePoint Context Detection (partially implemented)
- SharePoint Links Registry (sample data provided)
- Link Resolution Engine (✅ **COMPLETE**)
- Fuzzy Search Implementation (will use resolved links)
- Modal UI Component (will display resolved links)

## Files Created

```
src/
├── core/
│   └── link-resolver.ts        # Main resolution engine
├── utils/
│   └── context-detector.ts     # SharePoint context detection
├── types/
│   ├── sharepoint-context.ts   # Context type definitions
│   └── link-template.ts        # Link template types
├── data/
│   └── sample-links.ts         # Sample link templates
├── tests/
│   └── link-resolver.test.ts   # Test suite
└── index.ts                    # Public API exports
```

## Acceptance Criteria

- ✅ Placeholder replacement for `{webUrl}`, `{siteUrl}`, `{listId}`, `{listTitle}`, `{tenantUrl}`
- ✅ Link applicability logic (context-aware filtering by page type, version, required context)
- ✅ Link validation (unresolved placeholders, URL format, tenant verification)
- ✅ Final URL generation based on current context
- ✅ Comprehensive test suite demonstrating all features
- ✅ Type-safe implementation with TypeScript
- ✅ Modular architecture for easy extension

## Testing Checklist

- ✅ Placeholder replacement works correctly
- ✅ Multiple placeholders in one URL work correctly
- ✅ Links filter based on page type
- ✅ Links filter based on required context
- ✅ URL validation catches unresolved placeholders
- ✅ URL validation catches malformed URLs
- ✅ Context detection works for various SharePoint URLs
- ✅ Grouping by category works correctly
- ✅ Multiple link resolution works correctly
