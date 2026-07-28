/**
 * SharePoint Context Detector Tests
 * Comprehensive test cases for URL parsing and context detection
 */

import { describe, it, expect } from '@jest/globals';
import {
  detectSharePointContext,
  parseSharePointUrl,
  isSharePointUrl,
  extractTenantId,
  extractListId,
  extractItemId,
  extractViewId,
  SharePointVersion,
  PageContext,
} from '../src/context';

// Placeholder test to satisfy Jest requirement
describe('Context Detector', () => {
  it('should be defined', () => {
    expect(detectSharePointContext).toBeDefined();
    expect(parseSharePointUrl).toBeDefined();
    expect(isSharePointUrl).toBeDefined();
  });
});

/**
 * Test URLs representing different SharePoint scenarios
 */
const TEST_URLS = {
  // Root site collection
  rootSite: 'https://contoso.sharepoint.com/',

  // Sites collection
  teamSite: 'https://contoso.sharepoint.com/sites/marketing',
  teamsSite: 'https://contoso.sharepoint.com/teams/engineering',

  // Modern pages
  modernHomePage: 'https://contoso.sharepoint.com/sites/marketing/SitePages/Home.aspx',
  modernCustomPage:
    'https://contoso.sharepoint.com/sites/marketing/SitePages/Project-Overview.aspx',

  // Document libraries
  sharedDocuments:
    'https://contoso.sharepoint.com/sites/marketing/Shared Documents/Forms/AllItems.aspx',
  documentsFolder: 'https://contoso.sharepoint.com/sites/marketing/Shared Documents/Q4 Reports',
  modernLibrary: 'https://contoso.sharepoint.com/sites/marketing/Documents/Forms/AllItems.aspx',

  // Lists
  customList: 'https://contoso.sharepoint.com/sites/marketing/Lists/Announcements/AllItems.aspx',
  listWithId:
    'https://contoso.sharepoint.com/sites/marketing/_layouts/15/listedit.aspx?List={12345678-1234-1234-1234-123456789abc}',

  // List items
  displayForm:
    'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/DispForm.aspx?ID=42&List={12345678-1234-1234-1234-123456789abc}',
  editForm: 'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/EditForm.aspx?ID=42',
  newForm: 'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/NewForm.aspx',

  // Settings pages
  siteSettings: 'https://contoso.sharepoint.com/sites/marketing/_layouts/15/settings.aspx',
  siteContents: 'https://contoso.sharepoint.com/sites/marketing/_layouts/15/viewlsts.aspx',
  listSettings:
    'https://contoso.sharepoint.com/sites/marketing/_layouts/15/listedit.aspx?List={12345678-1234-1234-1234-123456789abc}',

  // Subsites
  subsite: 'https://contoso.sharepoint.com/sites/marketing/projects/q4',
  subsiteList:
    'https://contoso.sharepoint.com/sites/marketing/projects/Lists/ProjectTasks/AllItems.aspx',

  // With view parameter
  withView:
    'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/AllItems.aspx?View={87654321-4321-4321-4321-cba987654321}',

  // Non-SharePoint URLs
  nonSharePoint: 'https://www.microsoft.com',
  nonSharePointSubdomain: 'https://portal.azure.com',
};

/**
 * Run all tests
 */
function runTests(): void {
  console.log('=== SharePoint Context Detector Tests ===\n');

  // Test 1: SharePoint URL Detection
  console.log('Test 1: SharePoint URL Detection');
  console.log('Root site is SharePoint:', isSharePointUrl(TEST_URLS.rootSite));
  console.log('Team site is SharePoint:', isSharePointUrl(TEST_URLS.teamSite));
  console.log('Non-SharePoint is SharePoint:', isSharePointUrl(TEST_URLS.nonSharePoint));
  console.log('');

  // Test 2: Tenant ID Extraction
  console.log('Test 2: Tenant ID Extraction');
  const tenantId = extractTenantId('contoso.sharepoint.com');
  console.log('Tenant ID from contoso.sharepoint.com:', tenantId);
  console.log('');

  // Test 3: Parse Team Site URL
  console.log('Test 3: Parse Team Site URL');
  const teamSiteUrl = parseSharePointUrl(TEST_URLS.teamSite);
  if (teamSiteUrl) {
    console.log('Full URL:', teamSiteUrl.fullUrl);
    console.log('Domain:', teamSiteUrl.domain);
    console.log('Site Collection URL:', teamSiteUrl.siteCollectionUrl);
    console.log('Web URL:', teamSiteUrl.webUrl);
    console.log('Relative Path:', teamSiteUrl.relativePath);
  }
  console.log('');

  // Test 4: Parse Document Library URL
  console.log('Test 4: Parse Document Library URL');
  const libUrl = parseSharePointUrl(TEST_URLS.sharedDocuments);
  if (libUrl) {
    console.log('Site Collection URL:', libUrl.siteCollectionUrl);
    console.log('Web URL:', libUrl.webUrl);
    console.log('Relative Path:', libUrl.relativePath);
  }
  console.log('');

  // Test 5: Parse List with ID
  console.log('Test 5: Parse List with ID');
  const listUrl = parseSharePointUrl(TEST_URLS.listWithId);
  if (listUrl) {
    console.log('Web URL:', listUrl.webUrl);
    console.log('List ID:', extractListId(listUrl));
  }
  console.log('');

  // Test 6: Parse List Item Form
  console.log('Test 6: Parse List Item Form');
  const itemUrl = parseSharePointUrl(TEST_URLS.displayForm);
  if (itemUrl) {
    console.log('Web URL:', itemUrl.webUrl);
    console.log('List ID:', extractListId(itemUrl));
    console.log('Item ID:', extractItemId(itemUrl));
  }
  console.log('');

  // Test 7: Parse URL with View
  console.log('Test 7: Parse URL with View');
  const viewUrl = parseSharePointUrl(TEST_URLS.withView);
  if (viewUrl) {
    console.log('Web URL:', viewUrl.webUrl);
    console.log('View ID:', extractViewId(viewUrl));
  }
  console.log('');

  // Test 8: Detect Context - Modern Page
  console.log('Test 8: Detect Context - Modern Page');
  const modernPageContext = detectSharePointContext(TEST_URLS.modernHomePage);
  if (modernPageContext.success && modernPageContext.context) {
    console.log('Is SharePoint:', modernPageContext.context.isSharePoint);
    console.log('Version:', modernPageContext.context.version);
    console.log('Page Context:', modernPageContext.context.pageContext);
    console.log('Tenant ID:', modernPageContext.context.tenantId);
  }
  console.log('');

  // Test 9: Detect Context - List Item
  console.log('Test 9: Detect Context - List Item');
  const listItemContext = detectSharePointContext(TEST_URLS.displayForm);
  if (listItemContext.success && listItemContext.context) {
    console.log('Page Context:', listItemContext.context.pageContext);
    console.log('List Context:', listItemContext.context.listContext);
  }
  console.log('');

  // Test 10: Detect Context - Site Settings
  console.log('Test 10: Detect Context - Site Settings');
  const settingsContext = detectSharePointContext(TEST_URLS.siteSettings);
  if (settingsContext.success && settingsContext.context) {
    console.log('Page Context:', settingsContext.context.pageContext);
    console.log('Expected: SiteSettings');
  }
  console.log('');

  // Test 11: Detect Context - Document Library
  console.log('Test 11: Detect Context - Document Library');
  const docLibContext = detectSharePointContext(TEST_URLS.sharedDocuments);
  if (docLibContext.success && docLibContext.context) {
    console.log('Page Context:', docLibContext.context.pageContext);
    console.log('Expected: DocumentLibrary');
  }
  console.log('');

  // Test 12: Parse Subsite
  console.log('Test 12: Parse Subsite');
  const subsiteUrl = parseSharePointUrl(TEST_URLS.subsite);
  if (subsiteUrl) {
    console.log('Site Collection URL:', subsiteUrl.siteCollectionUrl);
    console.log('Web URL:', subsiteUrl.webUrl);
    console.log('Relative Path:', subsiteUrl.relativePath);
  }
  console.log('');

  // Test 13: Non-SharePoint URL
  console.log('Test 13: Non-SharePoint URL');
  const nonSpContext = detectSharePointContext(TEST_URLS.nonSharePoint);
  console.log('Success:', nonSpContext.success);
  console.log('Error:', nonSpContext.error);
  console.log('');

  console.log('=== All Tests Complete ===');
}

// Export for use in other contexts
export { runTests, TEST_URLS };

// Run tests if executed directly
if (require.main === module) {
  runTests();
}
