/**
 * SharePoint DevTools - Content Script
 * Implements modal UI component for quick navigation, column inspector, and object inspector
 */

import { Modal, ResultItem } from '../ui/modal';
import { ColumnInspectorIntegration } from './column-inspector-integration';
import { ObjectInspectorIntegration } from './object-inspector-integration';
import { PermissionInspectorIntegration } from './permission-inspector-integration';
import { LinkManager } from '../links/link-manager';
import { createFuzzySearch } from '../links/fuzzy-search';
import { PlaceholderValues, SharePointLink, LinkCategory } from '../links/types';
import {
  getSharePointContext,
  isSharePointPage,
  extractListIdAsync,
  setCachedListId,
} from '../context/sharepoint-context';
import { getSiteListsAndLibraries, SiteListInfo } from '../api/sharepoint-api';
import {
  getUserFriendlyErrorMessage,
  logError,
  PermissionError,
  AuthenticationError,
} from '../utils/error-handling';

console.log('SharePoint DevTools - Content Script loaded');

// Initialize modal and link manager
let quickNavModal: Modal | null = null;
let listsModal: Modal | null = null;
let linkManager: LinkManager | null = null;
let fuzzySearch: ReturnType<typeof createFuzzySearch> | null = null;
let listsFuzzySearch: ReturnType<typeof createFuzzySearch> | null = null;
let currentPlaceholders: PlaceholderValues = {};
let siteListsAndLibraries: SiteListInfo[] = [];
let siteListsFetched = false;

if (isSharePointPage()) {
  console.log('SharePoint page detected');

  // Initialize Column Inspector
  const columnInspector = new ColumnInspectorIntegration();
  columnInspector.init();

  // Initialize Object Inspector
  const objectInspector = new ObjectInspectorIntegration();
  objectInspector.init();

  // Initialize Permission Inspector
  const permissionInspector = new PermissionInspectorIntegration();
  permissionInspector.init();

  // Initialize LinkManager
  linkManager = new LinkManager();

  // Get current SharePoint context
  const context = getSharePointContext();
  currentPlaceholders = getCurrentPlaceholders();

  // Log context for debugging
  console.log('SharePoint Context:', {
    pageType: context.pageType,
    listId: context.listId,
    webUrl: context.webUrl,
    isModern: context.isModern,
  });
  console.log('Placeholders:', currentPlaceholders);

  // Get applicable links and create fuzzy search
  const isListContext = context.pageType === 'list' || context.pageType === 'library';
  const applicableLinks = linkManager.getApplicableLinks(currentPlaceholders, isListContext);
  fuzzySearch = createFuzzySearch(applicableLinks);

  // Detect dark mode preference
  const prefersDarkMode = window.matchMedia('(prefers-color-scheme: dark)').matches;

  // Create modal instance
  quickNavModal = new Modal({
    title: 'SharePoint DevTools',
    placeholder: 'Search for SharePoint locations...',
    darkMode: prefersDarkMode,
    onSearch: handleSearch,
    onSelect: handleSelect,
    onClose: handleClose,
  });

  // Create lists modal instance
  listsModal = new Modal({
    title: 'Lists & Libraries',
    placeholder: 'Search for lists and libraries...',
    darkMode: prefersDarkMode,
    onSearch: handleListsSearch,
    onSelect: handleSelect,
    onClose: handleClose,
  });

  console.log('QuickNav modal initialized with', applicableLinks.length, 'links');
  console.log('Lists modal initialized');

  // For classic SharePoint lists, try to fetch list ID asynchronously
  if (isListContext && !context.listId) {
    console.log('Attempting async list ID fetch for classic SharePoint...');
    fetchListIdAndUpdateContext();
  }

  // Start async fetch of site lists and libraries (don't block modal)
  fetchSiteListsAndLibraries();

  // Set up keyboard shortcut for Ctrl+K / Cmd+K (QuickNav)
  document.addEventListener('keydown', (e: KeyboardEvent) => {
    // Toggle QuickNav modal with Ctrl+K or Cmd+K
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
      e.preventDefault();
      if (quickNavModal) {
        quickNavModal.toggle();
      }
      return;
    }
  });

  console.log(
    'Keyboard shortcuts initialized: Ctrl+K (QuickNav), Alt+I (Object Inspector), Alt+P (Permissions)'
  );
}

/**
 * Asynchronously fetch list ID for classic SharePoint lists and update context
 */
async function fetchListIdAndUpdateContext(): Promise<void> {
  try {
    const listId = await extractListIdAsync();
    if (listId) {
      console.log('Successfully fetched list ID:', listId);

      // Cache the list ID so it's available for subsequent calls
      setCachedListId(listId);

      // Update placeholders with the fetched list ID
      currentPlaceholders = getCurrentPlaceholders();

      // Recreate fuzzy search with updated placeholders
      if (linkManager) {
        const context = getSharePointContext();
        const isListContext = context.pageType === 'list' || context.pageType === 'library';
        const applicableLinks = linkManager.getApplicableLinks(currentPlaceholders, isListContext);
        fuzzySearch = createFuzzySearch(applicableLinks);
        console.log(
          'Updated fuzzy search with',
          applicableLinks.length,
          'links (including list-specific)'
        );
      }
    }
  } catch (error) {
    logError('Error fetching list ID', error);
    // Don't show error to user - this is a non-blocking enhancement
  }
}

/**
 * Fetch site lists and libraries asynchronously
 * This runs in the background and doesn't block the main quicknav
 */
async function fetchSiteListsAndLibraries(): Promise<void> {
  const context = getSharePointContext();
  let webUrl = context.webUrl || window.location.origin;
  const siteUrl = context.siteUrl || window.location.origin;

  // Apply the same fix as getCurrentPlaceholders: if siteUrl has /sites/ but webUrl doesn't, use siteUrl
  // This ensures we get the correct site URL even when not on the site contents page
  if (siteUrl.includes('/sites/') && !webUrl.includes('/sites/')) {
    console.log('fetchSiteListsAndLibraries: webUrl is root, using siteUrl instead');
    webUrl = siteUrl;
  }

  console.log('Fetching site lists and libraries from:', webUrl);

  try {
    siteListsAndLibraries = await getSiteListsAndLibraries(webUrl);
    siteListsFetched = true;
    console.log('Site lists and libraries fetched:', siteListsAndLibraries.length, 'items');

    // Create fuzzy search index for lists and libraries
    const searchableItems = siteListsAndLibraries.map(listToSearchableItem);
    listsFuzzySearch = createFuzzySearch(searchableItems);
    console.log('Lists fuzzy search index created with', searchableItems.length, 'items');

    // If the lists modal is open, refresh the display
    if (listsModal && listsModal.isOpen()) {
      showAllLists();
    }
  } catch (error) {
    logError('Failed to fetch site lists and libraries', error);
    siteListsFetched = true; // Mark as fetched to show error instead of loading

    // If the lists modal is open, show the error
    if (listsModal && listsModal.isOpen()) {
      showListsError(error);
    }
  }
}

/**
 * Show error in the lists modal with user-friendly message
 */
function showListsError(error: unknown): void {
  if (!listsModal) return;

  const errorMessage = getUserFriendlyErrorMessage(error);
  const canRetry = !(error instanceof PermissionError || error instanceof AuthenticationError);

  listsModal.showError({
    title: 'Failed to Load Lists',
    message: errorMessage,
    canRetry,
    onRetry: () => {
      siteListsFetched = false;
      siteListsAndLibraries = [];
      fetchSiteListsAndLibraries();
      listsModal?.setState('loading', 'Retrying...');
    },
  });
}

/**
 * Get current SharePoint context placeholders
 */
function getCurrentPlaceholders(): PlaceholderValues {
  const context = getSharePointContext();

  // Sometimes webUrl returns the root instead of the site URL
  // If siteUrl has /sites/ but webUrl doesn't, use siteUrl for webUrl
  let webUrl = context.webUrl || window.location.origin;
  const siteUrl = context.siteUrl || window.location.origin;

  console.log('getCurrentPlaceholders - Before fix:', {
    webUrl,
    siteUrl,
    siteUrlHasSites: siteUrl.includes('/sites/'),
    webUrlHasSites: webUrl.includes('/sites/'),
  });

  if (siteUrl.includes('/sites/') && !webUrl.includes('/sites/')) {
    console.log('getCurrentPlaceholders: webUrl is root, using siteUrl instead');
    webUrl = siteUrl;
  }

  console.log('getCurrentPlaceholders - After fix:', { webUrl });

  // Extract tenant admin URL
  // From: https://contoso.sharepoint.com/sites/...
  // To:   https://contoso-admin.sharepoint.com
  let tenantAdminUrl = '';
  try {
    const url = new URL(window.location.origin);
    const hostname = url.hostname;

    // Check if it's a SharePoint Online URL
    if (hostname.includes('.sharepoint.com')) {
      // Extract tenant name (e.g., "contoso" from "contoso.sharepoint.com")
      const tenantName = hostname.split('.')[0];
      tenantAdminUrl = `https://${tenantName}-admin.sharepoint.com`;
    } else if (hostname.includes('.sharepoint-df.com')) {
      // Handle GCC/DoD environments
      const tenantName = hostname.split('.')[0];
      tenantAdminUrl = `https://${tenantName}-admin.sharepoint-df.com`;
    }
  } catch (error) {
    console.error('Error extracting tenant admin URL:', error);
  }

  const placeholders: PlaceholderValues = {
    webUrl: webUrl,
    siteUrl: siteUrl,
    serverUrl: window.location.origin,
    tenantAdminUrl: tenantAdminUrl,
  };

  // Add list-specific placeholders if in a list context
  if (context.listId) {
    placeholders.listId = context.listId;

    // If we have a listUrl from context, use it
    // Otherwise, use webUrl as fallback (works for most list settings pages)
    if (context.listUrl) {
      placeholders.listUrl = context.listUrl;
    } else {
      placeholders.listUrl = placeholders.webUrl;
    }
  }

  return placeholders;
}

/**
 * Convert SharePointLink to ResultItem
 */
function linkToResultItem(link: SharePointLink): ResultItem {
  const url = linkManager?.resolveLink(link, currentPlaceholders) || '#';

  // Log URL resolution for debugging
  if (link.id === 'site-contents' || link.id === 'site-settings' || link.id === 'list-settings') {
    console.log(`linkToResultItem [${link.id}]:`, {
      template: link.urlTemplate,
      resolvedUrl: url,
      placeholders: currentPlaceholders,
    });
  }

  return {
    id: link.id,
    title: link.title,
    description: link.description,
    url,
    category: link.category,
  };
}

/**
 * Convert SiteListInfo to ResultItem
 */
function listToResultItem(list: SiteListInfo): ResultItem {
  return {
    id: `site-list-${list.id}`,
    title: list.title,
    description: `${list.isLibrary ? 'Library' : 'List'} • ${list.itemCount} items`,
    url: list.url,
    category: 'Site Lists & Libraries',
  };
}

/**
 * Convert SiteListInfo to Fuse.js searchable object
 * Adds keywords for better fuzzy search matching
 */
function listToSearchableItem(list: SiteListInfo): SharePointLink {
  return {
    id: `site-list-${list.id}`,
    title: list.title,
    description: `${list.isLibrary ? 'Library' : 'List'} • ${list.itemCount} items`,
    category: LinkCategory.Content,
    keywords: [list.isLibrary ? 'library' : 'list', list.title],
    urlTemplate: '', // Not used for lists
    priority: 0,
  };
}

/**
 * Handle search query
 */
async function handleSearch(query: string): Promise<void> {
  console.log('Search query:', query);

  if (!quickNavModal || !fuzzySearch || !linkManager) return;

  if (query.trim() === '') {
    // Show all applicable links grouped by category (async for back-to-list feature)
    await showAllLinks();
  } else {
    // Use fuzzy search
    const results = fuzzySearch.search(query, { limit: 20 });

    if (results.length === 0) {
      quickNavModal.setState('idle');
      quickNavModal.setResults([]);
      return;
    }

    // Convert to ResultItems - show as flat list
    const resultItems = results.map((result) => linkToResultItem(result.item));

    quickNavModal.setResults(resultItems);
  }
}

/**
 * Check if current page is a list/library settings page
 */
function isListSettingsPage(): boolean {
  const url = window.location.href.toLowerCase();
  const settingsPatterns = [
    '/listedit.aspx',
    '/listgeneralsettings.aspx',
    '/list.aspx',
    '/advsetng.aspx', // Advanced settings
    '/versionsetup.aspx', // Versioning settings
    '/permissions.aspx', // Permissions page with ?obj=
    '/workflow.aspx', // Workflow settings
    '/contenttype.aspx', // Content type settings
    '/managedmetadata.aspx', // Managed metadata
  ];

  // Check if URL contains settings patterns and has a List parameter
  return (
    settingsPatterns.some((pattern) => url.includes(pattern)) &&
    (url.includes('list=') || url.includes('listid='))
  );
}

// Cache for list URL lookups
const listUrlCache = new Map<string, string>();

/**
 * Fetch list URL from SharePoint API using GUID
 */
async function fetchListUrlByGuid(webUrl: string, listGuid: string): Promise<string | undefined> {
  const cacheKey = `${webUrl}|${listGuid}`;

  // Check cache first
  if (listUrlCache.has(cacheKey)) {
    return listUrlCache.get(cacheKey);
  }

  try {
    const apiUrl = `${webUrl}/_api/web/lists(guid'${listGuid}')?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;

    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        Accept: 'application/json;odata=verbose',
      },
      credentials: 'same-origin',
    });

    if (response.ok) {
      const data = await response.json();
      const serverRelativeUrl = data.d?.RootFolder?.ServerRelativeUrl;

      if (serverRelativeUrl) {
        const fullUrl = `${new URL(webUrl).origin}${serverRelativeUrl}`;
        listUrlCache.set(cacheKey, fullUrl);
        return fullUrl;
      }
    }
  } catch (error) {
    console.error('Error fetching list URL:', error);
  }

  return undefined;
}

/**
 * Construct list view URL from list ID or context (async)
 */
async function getListViewUrl(
  context: ReturnType<typeof getSharePointContext>
): Promise<string | undefined> {
  // First try to get from context (this works for list pages, not settings pages)
  if (context.listUrl) {
    return context.listUrl;
  }

  // For settings pages, extract List GUID from URL parameters
  const urlParams = new URLSearchParams(window.location.search);
  const listParam = urlParams.get('List') || urlParams.get('ListId');

  if (listParam && context.webUrl) {
    // Extract just the GUID without curly braces
    const listGuid = listParam.replace(/[{}]/g, '').toLowerCase();

    // Fetch the actual list URL from API (includes internal name)
    const listUrl = await fetchListUrlByGuid(context.webUrl, listGuid);
    if (listUrl) {
      return listUrl;
    }
  }

  return undefined;
}

/**
 * Show default priority links (when search is empty)
 */
async function showAllLinks(): Promise<void> {
  if (!quickNavModal || !linkManager) return;

  const context = getSharePointContext();
  const isListContext = context.pageType === 'list' || context.pageType === 'library';

  // Refresh placeholders in case context has changed
  currentPlaceholders = getCurrentPlaceholders();

  console.log('showAllLinks - Context:', {
    pageType: context.pageType,
    isListContext,
    listId: context.listId,
  });
  console.log('showAllLinks - Placeholders:', currentPlaceholders);

  const applicableLinks = linkManager.getApplicableLinks(currentPlaceholders, isListContext);

  console.log('showAllLinks - Found', applicableLinks.length, 'applicable links');
  console.log('showAllLinks - Applicable link IDs:', applicableLinks.map((l) => l.id).slice(0, 10));

  // Filter to only show high-priority links (priority >= 5)
  // This gives us just the most important links instead of all 75+
  const priorityLinks = applicableLinks.filter((link) => (link.priority || 0) >= 5);

  console.log(
    'showAllLinks - Priority links (>=5):',
    priorityLinks.length,
    priorityLinks.map((l) => `${l.id}:${l.priority}`)
  );

  // Sort by priority (higher priority first), then alphabetically
  const sortedLinks = priorityLinks.sort((a, b) => {
    const priorityA = a.priority || 0;
    const priorityB = b.priority || 0;

    // If in list context, prioritize list-related links
    if (isListContext) {
      const aIsListLink = a.id.includes('list-') || a.id.includes('library-');
      const bIsListLink = b.id.includes('list-') || b.id.includes('library-');

      if (aIsListLink && !bIsListLink) return -1;
      if (!aIsListLink && bIsListLink) return 1;
    }

    // Sort by priority (descending)
    if (priorityB !== priorityA) {
      return priorityB - priorityA;
    }

    // Then alphabetically
    return a.title.localeCompare(b.title);
  });

  let resultItems = sortedLinks.map(linkToResultItem);

  // If on a list settings page, fetch and prepend "Back to List/Library" link at the top
  if (isListSettingsPage()) {
    const listViewUrl = await getListViewUrl(context);
    if (listViewUrl) {
      const backToListItem: ResultItem = {
        id: 'back-to-list',
        title: '← Back to List/Library',
        description: 'Return to the list or library view',
        url: listViewUrl,
        category: 'Navigation',
      };
      resultItems = [backToListItem, ...resultItems];
    }
  }

  // Don't group by category - show as flat list respecting priority order
  // This ensures Site Contents (priority 10) appears before Site Settings (priority 9)
  quickNavModal.setResults(resultItems);
}

/**
 * Show all lists and libraries in the lists modal
 */
function showAllLists(): void {
  if (!listsModal) return;

  if (siteListsFetched && siteListsAndLibraries.length > 0) {
    const resultItems = siteListsAndLibraries.map(listToResultItem);
    listsModal.setResults(resultItems);
    listsModal.setState('idle');
  } else if (siteListsFetched && siteListsAndLibraries.length === 0) {
    // Could be empty or could be error - check if we have an error state
    listsModal.setState('empty', 'No lists or libraries found');
  } else {
    listsModal.setState('loading', 'Loading lists and libraries...');
  }
}

/**
 * Handle search in the lists modal
 */
function handleListsSearch(query: string): void {
  if (!listsModal) return;

  if (!query.trim()) {
    showAllLists();
    return;
  }

  if (!siteListsFetched || siteListsAndLibraries.length === 0) {
    listsModal.setState('empty', 'No lists or libraries available');
    return;
  }

  // Use fuzzy search for lists and libraries
  if (!listsFuzzySearch) {
    console.warn('Lists fuzzy search not initialized');
    listsModal.setState('empty', 'Search index not ready');
    return;
  }

  const results = listsFuzzySearch.search(query, { limit: 50 });

  if (results.length > 0) {
    // Map fuzzy search results back to SiteListInfo items
    const resultItems = results
      .map((result) => {
        const list = siteListsAndLibraries.find((l) => `site-list-${l.id}` === result.item.id);
        return list ? listToResultItem(list) : null;
      })
      .filter((item) => item !== null) as ResultItem[];

    listsModal.setResults(resultItems);
    listsModal.setState('idle');
  } else {
    listsModal.setState('empty', `No lists or libraries match "${query}"`);
  }
}

/**
 * Handle result selection
 */
function handleSelect(item: ResultItem): void {
  console.log('Selected item:', item, 'URL:', item.url);

  if (item.url && item.url !== '#') {
    // Navigate to the selected URL
    window.location.href = item.url;
  }
}

/**
 * Handle modal close
 */
function handleClose(): void {
  console.log('Modal closed');
}

/**
 * Listen for messages from background script
 */
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'toggle-modal') {
    if (quickNavModal) {
      quickNavModal.toggle();
      sendResponse({ success: true });
    } else {
      sendResponse({ success: false, error: 'Modal not initialized' });
    }
  } else if (message.action === 'toggle-lists') {
    if (listsModal) {
      if (!listsModal.isOpen()) {
        showAllLists();
      }
      listsModal.toggle();
      sendResponse({ success: true });
    } else {
      sendResponse({ success: false, error: 'Lists modal not initialized' });
    }
  }
  return true;
});

console.log('Content script event listeners initialized');
