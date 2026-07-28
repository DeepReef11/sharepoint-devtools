# Fuzzy Search Implementation

## Overview

The SharePoint DevTools extension implements intelligent fuzzy search functionality using [Fuse.js](https://fusejs.io/), enabling users to quickly find SharePoint links even with typos, partial matches, or imprecise queries.

## Features

- **Typo Tolerance**: Finds matches even when search terms contain spelling errors
- **Multi-field Search**: Searches across title, description, keywords, and category
- **Relevance Ranking**: Results are scored and ranked by relevance
- **Configurable Matching**: Adjustable threshold and search parameters
- **Real-time Filtering**: Fast search performance for responsive UI
- **Partial Matching**: Supports incomplete search terms

## Architecture

### Components

1. **FuzzySearchService** (`src/links/fuzzy-search.ts`)
   - Core fuzzy search implementation
   - Wraps Fuse.js with SharePoint-specific configuration
   - Provides search, index management, and configuration methods

2. **LinkManager Integration** (`src/links/link-manager.ts`)
   - Integrates fuzzy search into the link management system
   - Provides convenience methods for different use cases
   - Maintains backward compatibility with simple search

3. **Type Definitions** (`src/links/fuzzy-search.ts`)
   - `SearchResult`: Search result with score and match information
   - `FuzzySearchOptions`: Configuration options for search queries

## Usage

### Basic Search

```typescript
import { linkManager } from './links';

// Perform fuzzy search
const results = linkManager.fuzzySearch('settings');

// Results include scores and matches
results.forEach(result => {
  console.log(result.item.title);      // Link title
  console.log(result.score);           // Match score (lower is better)
  console.log(result.matches);         // Field matches
});
```

### Search with Options

```typescript
// Limit results
const topResults = linkManager.fuzzySearch('permissions', {
  limit: 5
});

// Disable scores/matches in results
const simpleResults = linkManager.fuzzySearch('admin', {
  includeScore: false,
  includeMatches: false
});
```

### Get Links Only (No Scores)

```typescript
// Get array of SharePointLink objects
const links = linkManager.fuzzySearchLinks('recycle bin');

links.forEach(link => {
  console.log(link.title);
  console.log(link.description);
});
```

### Search a Subset of Links

```typescript
// Search only within specific links
const adminLinks = linkManager.getLinksByCategory(LinkCategory.SiteAdmin);
const results = linkManager.fuzzySearchSubset('features', adminLinks);
```

### Direct FuzzySearchService Usage

```typescript
import { createFuzzySearch } from './links/fuzzy-search';
import { SHAREPOINT_LINKS } from './links/registry';

// Create search service
const searchService = createFuzzySearch(SHAREPOINT_LINKS);

// Perform search
const results = searchService.search('site settings');

// Update index
searchService.setLinks(newLinks);

// Add/remove links
searchService.addLink(newLink);
searchService.removeLink('link-id');
```

## Configuration

### Default Fuse.js Options

```typescript
{
  threshold: 0.4,              // 0.0 = perfect, 1.0 = match anything
  location: 0,                 // Start of string
  distance: 100,               // How far to search
  minMatchCharLength: 2,       // Minimum characters to match
  ignoreLocation: true,        // Don't penalize matches far from start
  includeScore: true,          // Return match scores
  includeMatches: true,        // Return match metadata
}
```

### Field Weights

Fields are weighted by importance for ranking:

- **Title**: 2.0 (most important)
- **Keywords**: 1.5 (very relevant)
- **Description**: 1.0 (normal weight)
- **Category**: 0.8 (less important)

### Custom Configuration

```typescript
import { FuzzySearchService } from './links/fuzzy-search';

const searchService = new FuzzySearchService(links, {
  threshold: 0.3,      // More strict matching
  keys: [
    { name: 'title', weight: 3.0 },
    { name: 'keywords', weight: 2.0 }
  ]
});

// Or update options after creation
searchService.updateOptions({
  threshold: 0.5
});
```

## Search Examples

### Example 1: Typo Tolerance

```typescript
// Search with typo: "permisions" instead of "permissions"
const results = linkManager.fuzzySearch('permisions');
// ✓ Finds "Site Permissions"
```

### Example 2: Partial Matches

```typescript
// Search with partial term
const results = linkManager.fuzzySearch('recy');
// ✓ Finds "Recycle Bin"
```

### Example 3: Multi-word Search

```typescript
// Search with multiple words
const results = linkManager.fuzzySearch('site admin');
// ✓ Finds "Site Settings", "Site Admin" category links
```

### Example 4: Keyword Search

```typescript
// Search by keyword
const results = linkManager.fuzzySearch('security');
// ✓ Finds links with "security" in keywords
```

### Example 5: Category Search

```typescript
// Search by category
const results = linkManager.fuzzySearch('development');
// ✓ Finds links in "Development" category
```

## Search Result Format

```typescript
interface SearchResult {
  item: SharePointLink;           // The matched link
  score?: number;                 // Match score (0.0 = perfect)
  matches?: FuseResultMatch[];    // Match details
}

// Example result
{
  item: {
    id: 'site-settings',
    title: 'Site Settings',
    urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
    category: 'Site Admin',
    description: 'Main site settings page',
    keywords: ['settings', 'configuration', 'admin']
  },
  score: 0.001,
  matches: [
    {
      indices: [[0, 7]],
      value: 'Site Settings',
      key: 'title'
    }
  ]
}
```

## Performance Considerations

### Index Updates

The search index is automatically created when LinkManager is instantiated. If you modify links, update the index:

```typescript
linkManager.updateSearchIndex();
```

### Search Optimization

- **Minimum Query Length**: Queries under 2 characters return no results
- **Result Limiting**: Use `limit` option to cap results for better performance
- **Subset Search**: Search smaller link sets when appropriate

### Memory Usage

- Each FuzzySearchService instance maintains its own index
- Use the singleton `linkManager` instance for most cases
- Only create new instances when needed for different link sets

## Integration with UI

### Real-time Search Input

```typescript
// Debounce user input
let searchTimeout: number;
const searchInput = document.getElementById('search');

searchInput.addEventListener('input', (e) => {
  clearTimeout(searchTimeout);
  searchTimeout = setTimeout(() => {
    const query = e.target.value;
    const results = linkManager.fuzzySearch(query, { limit: 10 });
    displayResults(results);
  }, 300);
});
```

### Display Results with Scores

```typescript
function displayResults(results: SearchResult[]) {
  const list = document.getElementById('results');
  list.innerHTML = '';

  results.forEach(result => {
    const item = document.createElement('div');
    item.className = 'result-item';

    // Highlight the match score
    const relevance = (1 - (result.score || 0)) * 100;
    item.innerHTML = `
      <div class="title">${result.item.title}</div>
      <div class="description">${result.item.description}</div>
      <div class="category">${result.item.category}</div>
      <div class="relevance">${relevance.toFixed(0)}% match</div>
    `;

    list.appendChild(item);
  });
}
```

### Highlight Matched Text

```typescript
function highlightMatches(text: string, matches: FuseResultMatch[]): string {
  // Use matches.indices to highlight matched portions
  // Implementation depends on UI framework
}
```

## Backward Compatibility

The original `search()` method is still available but deprecated:

```typescript
// Old method (simple string matching)
const links = linkManager.search('settings');

// New method (fuzzy search)
const results = linkManager.fuzzySearch('settings');
const links = results.map(r => r.item);
```

## Testing

Run the fuzzy search test suite:

```typescript
import { runAllFuzzySearchTests } from './tests/fuzzy-search.test';

// In browser console or Node.js
runAllFuzzySearchTests();

// Or run individual tests
import { testFuzzyMatching } from './tests/fuzzy-search.test';
testFuzzyMatching();
```

## Best Practices

1. **Use fuzzySearch for User Input**: Provides better UX with typo tolerance
2. **Set Appropriate Limits**: Cap results at 10-20 for UI display
3. **Combine with Context Filtering**: Search within applicable links only
4. **Display Relevance Scores**: Help users understand match quality
5. **Debounce Input**: Avoid excessive searches while typing
6. **Cache Results**: Store recent searches if appropriate

## Future Enhancements

Potential improvements for future versions:

- Search history and suggestions
- Custom scoring algorithms
- Phonetic matching
- Synonym support
- Search analytics
- Personalized ranking based on usage

## References

- [Fuse.js Documentation](https://fusejs.io/)
- [Fuse.js GitHub](https://github.com/krisk/fuse)
- [SharePoint DevTools Roadmap](../ROADMAP.md)
