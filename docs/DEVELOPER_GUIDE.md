# SharePoint DevTools - Developer Guide

## Table of Contents
1. [Architecture Overview](#architecture-overview)
2. [Project Structure](#project-structure)
3. [Development Setup](#development-setup)
4. [Build System](#build-system)
5. [Core Components](#core-components)
6. [Testing](#testing)
7. [Code Style and Standards](#code-style-and-standards)
8. [Contributing Guidelines](#contributing-guidelines)
9. [Release Process](#release-process)
10. [Troubleshooting Development Issues](#troubleshooting-development-issues)

## Architecture Overview

SharePoint DevTools is a Chrome/Edge browser extension built with:
- **Language**: TypeScript
- **Build Tool**: Webpack
- **Extension Platform**: Chrome Extension Manifest V3
- **Testing**: Jest with ts-jest
- **Fuzzy Search**: Fuse.js
- **UI**: Vanilla TypeScript (no framework dependency)

### Design Principles
1. **Lightweight**: No heavy frameworks, minimal dependencies
2. **Performance**: Lazy loading, efficient caching
3. **Type Safety**: Comprehensive TypeScript types
4. **Modularity**: Clear separation of concerns
5. **Testability**: Unit and integration tests

### Extension Components
```
┌─────────────────────────────────────────────┐
│           Browser Extension                  │
├─────────────────────────────────────────────┤
│  Content Scripts  │  Background Worker      │
│  (SharePoint Pages) │  (Service Worker)    │
├─────────────────┬─────────────────────────┤
│                 │                            │
│  UI Components  │  Storage & State         │
│  - QuickNav Modal│  - Chrome Storage API    │
│  - Column Inspector│ - Recent/Favorites     │
│  - Object Inspector│ - Custom Links         │
│                 │                            │
├─────────────────┴──────────────────────────┤
│           Core Modules                       │
│  - Context Detection                         │
│  - Link Registry & Resolution                │
│  - Fuzzy Search                              │
│  - SharePoint API Client                     │
└─────────────────────────────────────────────┘
```

## Project Structure

```
sharepoint-devtools/
├── src/
│   ├── api/                    # SharePoint REST API integrations
│   │   ├── sharepoint-api.ts   # Main API client
│   │   ├── column-fetcher.ts   # Column metadata fetching
│   │   └── api-cache.ts        # Response caching
│   │
│   ├── background/             # Background service worker
│   │   ├── service-worker.ts   # Manifest V3 service worker
│   │   └── background.ts       # Background logic
│   │
│   ├── content/                # Content scripts (run on SharePoint pages)
│   │   ├── content-script.ts   # Main content script entry
│   │   ├── content.ts          # Legacy content script
│   │   ├── modal.ts            # Modal management
│   │   ├── column-inspector-integration.ts
│   │   └── object-inspector-integration.ts
│   │
│   ├── context/                # SharePoint context detection
│   │   ├── sharepoint-context.ts    # Context extraction
│   │   ├── context-detector.ts      # Detection logic
│   │   └── url-parser.ts           # URL parsing utilities
│   │
│   ├── links/                  # Link registry and search
│   │   ├── registry.ts         # Link definitions
│   │   ├── link-manager.ts     # Link management
│   │   ├── fuzzy-search.ts     # Fuse.js integration
│   │   ├── template-engine.ts  # URL template resolution
│   │   └── types.ts            # Link type definitions
│   │
│   ├── managers/               # Business logic managers
│   │   └── tenant-manager.ts   # Multi-tenant management
│   │
│   ├── storage/                # Storage abstractions
│   │   ├── storage-manager.ts  # Chrome storage wrapper
│   │   ├── favorites-manager.ts # Favorites persistence
│   │   ├── link-tracker.ts     # Recent links tracking
│   │   ├── tenant-storage.ts   # Tenant data storage
│   │   └── custom-links-storage.ts
│   │
│   ├── types/                  # TypeScript type definitions
│   │   ├── sharepoint.ts       # SharePoint types
│   │   ├── sharepoint-context.ts
│   │   ├── messages.ts         # Extension messaging types
│   │   ├── link-template.ts
│   │   └── column-schema.ts
│   │
│   ├── ui/                     # UI components
│   │   ├── modal/              # Base modal component
│   │   │   ├── Modal.ts
│   │   │   └── types.ts
│   │   ├── inspector/          # Inspector modal
│   │   │   └── InspectorModal.ts
│   │   ├── column-inspector.ts
│   │   ├── object-inspector.ts
│   │   └── quick-access.ts
│   │
│   ├── utils/                  # Utility functions
│   │   ├── sharepoint-detector.ts
│   │   ├── context-detector.ts
│   │   ├── error-handling.ts
│   │   └── clipboard.ts
│   │
│   ├── options/                # Extension options page
│   │   └── options.ts
│   │
│   └── index.ts                # Main entry point
│
├── tests/                      # Test files
│   ├── setup.ts                # Jest setup
│   ├── fuzzy-search.test.ts
│   ├── clipboard.test.ts
│   ├── sharepoint-detector.test.ts
│   └── error-handling.test.ts
│
├── public/                     # Static assets
│   ├── manifest.json           # Extension manifest
│   ├── icons/                  # Extension icons
│   └── options.html            # Options page HTML
│
├── docs/                       # Documentation
│   ├── USER_GUIDE.md
│   └── DEVELOPER_GUIDE.md
│
├── webpack.*.js                # Webpack configuration
├── tsconfig.json               # TypeScript configuration
├── jest.config.js              # Jest configuration
├── package.json                # Dependencies and scripts
└── README.md                   # Project overview
```

## Development Setup

### Prerequisites
- **Node.js**: v18 or higher
- **npm**: v8 or higher (comes with Node.js)
- **Git**: For version control
- **Browser**: Chrome or Edge for testing

### Initial Setup
```bash
# Clone the repository
git clone https://github.com/DeepReef11/sharepoint-devtools.git
cd sharepoint-devtools

# Install dependencies
npm install

# Build the extension
npm run build:dev

# Or watch for changes
npm run watch
```

### Loading the Extension in Browser
1. Build the extension (see above)
2. Open Chrome/Edge: `chrome://extensions/` or `edge://extensions/`
3. Enable "Developer mode" (toggle in top-right)
4. Click "Load unpacked"
5. Select the `dist` folder
6. Test on a SharePoint site

### Development Workflow
```bash
# Start watch mode (auto-rebuild on file changes)
npm run watch

# In another terminal, reload extension in browser after changes
# or use Extension Reload extension for auto-reload
```

## Build System

### Webpack Configuration
The project uses Webpack with three configurations:

#### webpack.common.js
Common configuration shared across all builds:
- TypeScript loader (ts-loader)
- Multiple entry points (content script, background, options)
- Asset management (icons, HTML)
- Source maps

#### webpack.dev.js
Development build configuration:
- Faster build times (no minification)
- Source maps for debugging
- Watch mode compatible

#### webpack.prod.js
Production build configuration:
- Minification enabled
- Optimized bundle size
- Clean output directory

### Build Scripts
```bash
# Development build (fast, with source maps)
npm run build:dev

# Production build (optimized, minified)
npm run build

# Watch mode (auto-rebuild on changes)
npm run watch

# Type checking (no build)
npm run type-check

# Linting
npm run lint

# Auto-fix linting issues
npm run lint:fix

# Format code with Prettier
npm run format

# Check formatting
npm run format:check
```

### Output Structure
```
dist/
├── manifest.json
├── content-script.js
├── background.js
├── options.js
├── options.html
├── icons/
└── *.js.map (in dev mode)
```

## Core Components

### 1. Context Detection (`src/context/`)

**Purpose**: Detect and extract SharePoint context from the current page.

**Key Functions**:
```typescript
// Detect if page is SharePoint
isSharePointPage(): boolean

// Extract site collection URL
extractSiteUrl(): string | undefined

// Extract list GUID (async version for API fallback)
extractListIdAsync(): Promise<string | undefined>

// Get full context object
getSharePointContext(): SharePointContext
```

**Usage Example**:
```typescript
import { getSharePointContext } from './context/sharepoint-context';

const context = getSharePointContext();
if (context.isSharePoint && context.listId) {
  // We're on a list page, listId is available
}
```

### 2. Link Registry (`src/links/`)

**Purpose**: Define and manage navigable SharePoint links.

**Link Definition**:
```typescript
interface SharePointLink {
  id: string;                    // Unique identifier
  title: string;                 // Display name
  urlTemplate: string;           // URL with placeholders
  category: LinkCategory;        // Organization category
  description?: string;          // Search description
  keywords?: string[];           // Search keywords
  requiredContext?: string[];    // Required context (e.g., ['listId'])
  priority?: number;             // Display priority
}
```

**URL Template Placeholders**:
- `{siteUrl}`: Site collection URL
- `{webUrl}`: Web URL (subsite aware)
- `{listId}`: List GUID
- `{listUrl}`: List URL
- `{tenant}`: Tenant name

**Adding New Links**:
Edit `src/links/registry.ts`:
```typescript
export const SHAREPOINT_LINKS: SharePointLink[] = [
  {
    id: 'my-custom-link',
    title: 'My Custom Page',
    urlTemplate: '{webUrl}/_layouts/15/mycustom.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'Access my custom page',
    keywords: ['custom', 'my'],
    priority: 1,
  },
  // ... existing links
];
```

### 3. Fuzzy Search (`src/links/fuzzy-search.ts`)

**Purpose**: Provide fuzzy searching using Fuse.js.

**Configuration**:
```typescript
const FUSE_OPTIONS = {
  threshold: 0.4,        // 0.0 = exact, 1.0 = match anything
  keys: [
    { name: 'title', weight: 2.0 },
    { name: 'keywords', weight: 1.5 },
    { name: 'description', weight: 1.0 },
    { name: 'category', weight: 0.8 },
  ],
};
```

**Usage**:
```typescript
import { createFuzzySearch } from './links/fuzzy-search';

const searchService = createFuzzySearch(links);
const results = searchService.search('settings');
// Returns: SearchResult[] with scores
```

### 4. SharePoint API Client (`src/api/sharepoint-api.ts`)

**Purpose**: Interact with SharePoint REST API.

**Key Methods**:
```typescript
class SharePointAPI {
  // Get list columns
  async getListColumns(webUrl: string, listId: string): Promise<Column[]>

  // Get site properties
  async getSiteProperties(siteUrl: string): Promise<SiteProperties>

  // Get all lists
  async getLists(webUrl: string): Promise<List[]>

  // Get content types
  async getContentTypes(webUrl: string): Promise<ContentType[]>
}
```

**API Caching**:
The API client includes automatic caching:
```typescript
// Cache configuration
const CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

// Cached requests
const columns = await api.getListColumns(webUrl, listId);
// Subsequent calls within 5 minutes use cache
```

### 5. Storage Management (`src/storage/`)

**Purpose**: Persist user data (favorites, recent links, custom links).

**Storage Locations**:
- `chrome.storage.sync`: Custom links (synced across devices)
- `chrome.storage.local`: Recent links, favorites (device-specific)

**Key Managers**:
```typescript
// Favorites
import { FavoritesManager } from './storage/favorites-manager';
await FavoritesManager.addFavorite(linkId);
await FavoritesManager.removeFavorite(linkId);
const favorites = await FavoritesManager.getFavorites();

// Recent Links
import { LinkTracker } from './storage/link-tracker';
await LinkTracker.trackLinkAccess(linkId, linkData);
const recent = await LinkTracker.getRecentLinks(10);

// Custom Links
import { CustomLinksStorage } from './storage/custom-links-storage';
await CustomLinksStorage.saveCustomLink(customLink);
const customLinks = await CustomLinksStorage.getCustomLinks();
```

### 6. UI Components (`src/ui/`)

**Base Modal Component**:
```typescript
import { Modal } from './ui/modal/Modal';

class MyModal extends Modal {
  constructor(options: ModalOptions) {
    super(options);
  }

  protected renderContent(): string {
    return `<div>My modal content</div>`;
  }

  protected setupEventListeners(): void {
    // Add event listeners
  }
}
```

**Inspector Modal**:
Pre-built component for displaying metadata:
```typescript
import { InspectorModal } from './ui/inspector/InspectorModal';

const modal = new InspectorModal({
  title: 'Column Inspector',
  data: columns,
  onClose: () => console.log('Closed'),
});

modal.show();
```

## Testing

### Running Tests
```bash
# Run all tests
npm test

# Watch mode (re-run on changes)
npm run test:watch

# Coverage report
npm run test:coverage

# Verbose output
npm run test:verbose
```

### Writing Tests

**Unit Test Example** (`src/utils/__tests__/clipboard.test.ts`):
```typescript
import { describe, it, expect, jest } from '@jest/globals';
import { copyToClipboard } from '../clipboard';

describe('copyToClipboard', () => {
  it('should copy text using clipboard API', async () => {
    Object.assign(navigator, {
      clipboard: {
        writeText: jest.fn().mockResolvedValue(undefined),
      },
    });

    const result = await copyToClipboard('test');
    expect(result).toBe(true);
    expect(navigator.clipboard.writeText).toHaveBeenCalledWith('test');
  });
});
```

**Integration Test Example**:
```typescript
describe('SharePointContext Integration', () => {
  it('should extract context from list page URL', () => {
    // Mock window.location
    Object.defineProperty(window, 'location', {
      value: {
        href: 'https://contoso.sharepoint.com/sites/test/Lists/Tasks/AllItems.aspx?List={guid}',
        pathname: '/sites/test/Lists/Tasks/AllItems.aspx',
        search: '?List={guid}',
      },
      writable: true,
    });

    const context = getSharePointContext();
    expect(context.isSharePoint).toBe(true);
    expect(context.listId).toBeDefined();
    expect(context.pageType).toBe('list');
  });
});
```

### Test Coverage
Target: **>70% code coverage**

```bash
# Generate coverage report
npm run test:coverage

# View HTML report
open coverage/index.html
```

### Mocking Chrome APIs
Jest setup (`tests/setup.ts`) includes Chrome API mocks:
```typescript
(global as any).chrome = {
  runtime: {
    sendMessage: jest.fn(),
    onMessage: {
      addListener: jest.fn(),
    },
  },
  storage: {
    sync: {
      get: jest.fn(),
      set: jest.fn(),
    },
  },
};
```

## Code Style and Standards

### TypeScript Guidelines
1. **Strict Type Safety**: Use explicit types, avoid `any`
2. **Interfaces Over Types**: Prefer interfaces for object shapes
3. **Named Exports**: Use named exports over default exports
4. **Async/Await**: Prefer async/await over promises
5. **JSDoc Comments**: Document public functions and classes

**Example**:
```typescript
/**
 * Extracts the site URL from a SharePoint page
 * @param url - The URL to parse (defaults to current page)
 * @returns The site collection URL or undefined if not found
 */
export function extractSiteUrl(url: string = window.location.href): string | undefined {
  const match = url.match(/(https?:\/\/[^/]+\/sites\/[^/]+)/);
  return match ? match[1] : undefined;
}
```

### File Organization
- **One class/module per file**
- **Related functionality grouped in directories**
- **Index files for clean imports** (`src/links/index.ts`)

### Naming Conventions
- **Files**: `kebab-case.ts`
- **Classes**: `PascalCase`
- **Functions**: `camelCase`
- **Constants**: `UPPER_SNAKE_CASE`
- **Interfaces**: `PascalCase` (no `I` prefix)
- **Types**: `PascalCase`

### Code Formatting
- **Tool**: Prettier
- **Config**: `.prettierrc`
- **Integration**: ESLint + Prettier

```bash
# Format all files
npm run format

# Check formatting
npm run format:check
```

### Linting
- **Tool**: ESLint with TypeScript plugin
- **Config**: `eslint.config.mjs`

```bash
# Lint code
npm run lint

# Auto-fix issues
npm run lint:fix
```

## Contributing Guidelines

### Before Starting
1. **Check existing issues**: Avoid duplicate work
2. **Create an issue**: Describe your proposed change
3. **Get feedback**: Discuss approach before coding
4. **Fork the repository**: Work in your own fork

### Development Process
1. **Create a feature branch**:
   ```bash
   git checkout -b feature/my-new-feature
   ```

2. **Make your changes**:
   - Follow code style guidelines
   - Add tests for new functionality
   - Update documentation
   - Ensure tests pass: `npm test`
   - Ensure linting passes: `npm run lint`

3. **Commit your changes**:
   ```bash
   git add .
   git commit -m "Add feature: description of feature"
   ```

   **Commit Message Format**:
   ```
   <type>: <subject>

   <body>

   <footer>
   ```

   **Types**: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`

4. **Push to your fork**:
   ```bash
   git push origin feature/my-new-feature
   ```

5. **Create a Pull Request**:
   - Clear title and description
   - Link to related issues
   - Include screenshots for UI changes
   - Ensure CI passes

### Pull Request Checklist
- [ ] Code follows style guidelines
- [ ] Self-reviewed code
- [ ] Commented hard-to-understand areas
- [ ] Updated documentation
- [ ] Added tests
- [ ] All tests passing
- [ ] No new linting errors
- [ ] Tested in browser (Chrome and Edge if possible)

### Code Review Process
1. Maintainer reviews PR
2. Requested changes are addressed
3. PR is approved
4. PR is merged to main

## Release Process

### Versioning
We follow [Semantic Versioning](https://semver.org/):
- **Major**: Breaking changes (e.g., 1.0.0 → 2.0.0)
- **Minor**: New features (e.g., 1.0.0 → 1.1.0)
- **Patch**: Bug fixes (e.g., 1.0.0 → 1.0.1)

### Release Steps
1. **Update version** in `package.json` and `manifest.json`
2. **Update CHANGELOG.md** with changes
3. **Run full test suite**: `npm test`
4. **Build production**: `npm run build`
5. **Test the build** in browser
6. **Create git tag**:
   ```bash
   git tag -a v1.0.0 -m "Release version 1.0.0"
   git push origin v1.0.0
   ```
7. **Create GitHub release** with changelog
8. **Package extension**:
   ```bash
   cd dist
   zip -r ../sharepoint-devtools-v1.0.0.zip *
   ```
9. **Upload to Chrome Web Store / Edge Add-ons**

## Troubleshooting Development Issues

### Build Failures

**Issue**: Webpack build fails
```bash
# Clean and rebuild
rm -rf dist node_modules package-lock.json
npm install
npm run build
```

**Issue**: TypeScript errors
```bash
# Check TypeScript version
npm list typescript

# Run type checker
npm run type-check
```

### Extension Not Loading

**Issue**: Extension won't load in browser
- Check for errors in `chrome://extensions/`
- Verify `manifest.json` is valid
- Ensure all referenced files exist in `dist/`

### Hot Reload Not Working

**Issue**: Changes not appearing
- Hard refresh: Ctrl+Shift+R
- Reload extension manually in `chrome://extensions/`
- Check watch mode is running: `npm run watch`

### Test Failures

**Issue**: Tests failing locally
```bash
# Clear Jest cache
npm test -- --clearCache

# Run specific test
npm test -- fuzzy-search.test.ts

# Debug mode
node --inspect-brk node_modules/.bin/jest --runInBand
```

### API Errors

**Issue**: SharePoint API calls failing
- Check browser console (F12) for CORS errors
- Verify you're logged into SharePoint
- Check API endpoint URL is correct
- Verify list/site GUIDs are valid

### Performance Issues

**Issue**: Extension running slowly
- Check for memory leaks in DevTools
- Profile with Chrome DevTools Performance tab
- Review caching strategy
- Optimize fuzzy search configuration

## Additional Resources

### Useful Links
- [Chrome Extension Docs](https://developer.chrome.com/docs/extensions/)
- [SharePoint REST API Reference](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [Fuse.js Documentation](https://fusejs.io/)
- [TypeScript Handbook](https://www.typescriptlang.org/docs/handbook/intro.html)
- [Jest Documentation](https://jestjs.io/docs/getting-started)

### Extension Development Tools
- **Extension Reloader**: Auto-reload extension on file changes
- **Chrome DevTools**: Debug content scripts and background workers
- **React DevTools**: Inspect React components (if using React)

### Community
- **GitHub Issues**: Report bugs, request features
- **GitHub Discussions**: Ask questions, share ideas

---

**Version**: 1.0.0
**Maintainer**: DeepReef11
**License**: MIT
