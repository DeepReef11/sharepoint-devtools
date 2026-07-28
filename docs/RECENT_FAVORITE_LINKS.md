# Recent and Favorite Links Feature

## Overview

The Recent and Favorite Links feature enhances SharePoint DevTools by providing users with quick access to their most frequently used links and personally curated favorites. This feature uses `chrome.storage` to persist user data across sessions.

## Features

### Recent Links
- **Automatic Tracking**: Links are automatically tracked when accessed
- **Access Count**: Tracks how many times each link has been accessed
- **Smart Sorting**: Most recently accessed links appear first
- **Configurable Limit**: Default is 10 recent links (configurable via preferences)

### Favorite Links
- **Manual Selection**: Users can mark any link as a favorite
- **Persistent Storage**: Favorites are saved across browser sessions
- **Custom Notes**: Add optional notes to favorite links
- **Flexible Sorting**: Sort by recent, alphabetical, or category

### Quick Access Section
- **Integrated UI**: Displays recent and favorite links in an easy-to-access section
- **One-Click Access**: Click any link to navigate
- **Toggle Favorites**: Quickly add/remove favorites with star icons
- **Dark Mode Support**: Automatically adapts to user's color scheme preference

## Architecture

### Storage Structure

The feature uses three main storage keys in `chrome.storage.local`:

1. **`recentLinks`**: Stores recently accessed links
   ```typescript
   {
     links: RecentLink[],
     maxSize: number
   }
   ```

2. **`favoriteLinks`**: Stores user's favorite links
   ```typescript
   {
     links: { [linkId: string]: FavoriteLink }
   }
   ```

3. **`linkPreferences`**: User preferences
   ```typescript
   {
     maxRecentLinks: number,
     showRecent: boolean,
     showFavorites: boolean,
     favoriteSortOrder: 'recent' | 'alphabetical' | 'category'
   }
   ```

### Components

#### Storage Layer (`src/storage/`)

- **`storage-manager.ts`**: Core storage operations with chrome.storage API
  - Get/save recent links
  - Get/save favorite links
  - Manage user preferences
  - Export/import functionality

- **`link-tracker.ts`**: Tracks link access and manages recent history
  - Track link clicks
  - Get recent links
  - Get most frequently accessed links
  - Clear history

- **`favorites-manager.ts`**: Manages favorite links
  - Add/remove favorites
  - Toggle favorite status
  - Sort and group favorites
  - Export/import favorites

#### UI Layer (`src/ui/`)

- **`quick-access.ts`**: Quick access UI component
  - Renders recent and favorite links
  - Handles user interactions
  - Responsive design with dark mode support

#### Integration Layer (`src/links/`)

- **`link-manager.ts`**: Enhanced with tracking methods
  - Track link access
  - Get recent links
  - Manage favorites
  - Integrated with existing link resolution

## Usage

### Basic Usage

```typescript
import { LinkTracker, FavoritesManager } from './storage';
import { linkManager } from './links/link-manager';

// Track a link access
const link = linkManager.getLinkById('site-settings');
const resolvedUrl = linkManager.resolveLink(link, placeholderValues);
await LinkTracker.trackLinkAccess(link, resolvedUrl);

// Add a link to favorites
await FavoritesManager.addFavorite(link, 'My favorite settings page');

// Check if a link is favorited
const isFav = await FavoritesManager.isFavorite('site-settings');

// Get recent links
const recentLinks = await LinkTracker.getRecentLinks(5);

// Get all favorites
const favorites = await FavoritesManager.getFavorites();
```

### Using the Quick Access Component

```typescript
import { QuickAccess } from './ui/quick-access';

// Initialize Quick Access component
const quickAccess = new QuickAccess({
  onLinkClick: async (linkId, url) => {
    // Custom link click handler
    window.open(url, '_blank');
  },
  placeholderValues: {
    webUrl: 'https://example.sharepoint.com/sites/mysite',
    siteUrl: 'https://example.sharepoint.com/sites/mysite',
  },
});

// Render in a container
const container = document.getElementById('quick-access-container');
await quickAccess.render(container);

// Inject styles
QuickAccess.injectStyles();
```

### Using with Link Manager

```typescript
import { linkManager } from './links/link-manager';

// Get recent links through link manager
const recentLinks = await linkManager.getRecentLinks(10);

// Get favorites through link manager
const favorites = await linkManager.getFavorites();

// Toggle favorite status
const link = linkManager.getLinkById('site-settings');
const isFavorited = await linkManager.toggleFavorite(link);

// Track link access
const resolvedUrl = linkManager.resolveLink(link, placeholderValues);
await linkManager.trackLinkAccess(link, resolvedUrl);
```

## User Preferences

Users can customize the behavior through preferences:

```typescript
import { StorageManager } from './storage';

// Get current preferences
const prefs = await StorageManager.getPreferences();

// Update preferences
await StorageManager.savePreferences({
  maxRecentLinks: 20,        // Increase recent links limit
  showRecent: true,           // Show recent section
  showFavorites: true,        // Show favorites section
  favoriteSortOrder: 'alphabetical'  // Sort favorites alphabetically
});
```

## Data Management

### Export Data

```typescript
import { StorageManager } from './storage';

// Export all data
const data = await StorageManager.exportData();
console.log(JSON.stringify(data, null, 2));

// Export only favorites
const favoritesJson = await FavoritesManager.exportFavorites();
```

### Import Data

```typescript
import { StorageManager } from './storage';

// Import data
await StorageManager.importData({
  recent: [...],
  favorites: [...],
  preferences: {...}
});

// Import favorites from JSON
const count = await FavoritesManager.importFavorites(jsonString);
console.log(`Imported ${count} favorites`);
```

### Clear Data

```typescript
import { LinkTracker, FavoritesManager } from './storage';

// Clear recent links
await LinkTracker.clearHistory();

// Clear all favorites
await FavoritesManager.clearAll();
```

## Storage Limits

Chrome extensions have storage limits for `chrome.storage.local`:

- **Default Quota**: 5MB for most extensions
- **Unlimited Storage**: Can be requested via manifest permission (not currently enabled)

### Current Implementation Limits

- **Recent Links**: Default maximum of 10 links (configurable)
- **Favorites**: No hard limit, but recommended to keep under 100 for performance
- **Storage Monitoring**: Use `StorageManager.getStorageStats()` to monitor usage

```typescript
const stats = await StorageManager.getStorageStats();
console.log(`Recent: ${stats.recentCount}, Favorites: ${stats.favoriteCount}`);
console.log(`Storage used: ${stats.bytesUsed} bytes`);
```

## Performance Considerations

1. **Async Operations**: All storage operations are asynchronous
2. **Batching**: Consider batching multiple reads/writes when possible
3. **Caching**: UI components should cache data when appropriate
4. **Lazy Loading**: Load data only when needed

## Testing

The feature includes comprehensive unit tests:

- `storage-manager.test.ts`: Tests for storage operations
- `link-tracker.test.ts`: Tests for link tracking
- Additional tests can be added for favorites-manager and UI components

Run tests with:
```bash
npm test
```

## Future Enhancements

Potential improvements for future versions:

1. **Sync Across Devices**: Use `chrome.storage.sync` for cross-device synchronization
2. **Advanced Analytics**: Track additional metrics (time spent, navigation patterns)
3. **Smart Suggestions**: AI-powered link recommendations based on usage patterns
4. **Collections**: Group favorites into custom collections/folders
5. **Keyboard Shortcuts**: Quick access to recent/favorite links via keyboard
6. **Search**: Search within recent and favorite links
7. **Bulk Operations**: Bulk add/remove favorites, bulk export/import

## Troubleshooting

### Common Issues

**Issue**: Recent links not being tracked
- **Solution**: Ensure `trackLinkAccess()` is called after link navigation
- **Check**: Verify storage permissions in manifest.json

**Issue**: Favorites not persisting
- **Solution**: Check browser's storage quota and clear if necessary
- **Check**: Use browser DevTools > Application > Storage to inspect

**Issue**: UI not updating after changes
- **Solution**: Ensure component re-renders after storage operations
- **Check**: Verify event listeners are properly attached

### Debug Mode

Enable debug logging:

```typescript
// Add this to enable verbose logging
localStorage.setItem('sp-quicknav-debug', 'true');
```

## API Reference

See inline TypeScript documentation for detailed API reference:

- [StorageManager](../src/storage/storage-manager.ts)
- [LinkTracker](../src/storage/link-tracker.ts)
- [FavoritesManager](../src/storage/favorites-manager.ts)
- [QuickAccess](../src/ui/quick-access.ts)

## Contributing

When adding new features to the Recent/Favorite Links system:

1. **Update Types**: Add new types to `src/storage/types.ts`
2. **Add Tests**: Write unit tests for new functionality
3. **Update Docs**: Update this documentation
4. **Consider Storage**: Be mindful of storage quota usage
5. **Maintain Backwards Compatibility**: Handle migration of existing data

## License

This feature is part of SharePoint DevTools and follows the project's license.
