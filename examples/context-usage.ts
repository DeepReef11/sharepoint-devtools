/**
 * SharePoint Context Detection - Usage Examples
 * Demonstrates how to use the context detection module in various scenarios
 */

import {
  detectSharePointContext,
  detectCurrentContext,
  parseSharePointUrl,
  isSharePointUrl,
  SharePointContext,
  PageContext,
  SharePointVersion,
} from '../src/context';

/**
 * Example 1: Basic URL Validation
 */
function example1_validateSharePointUrl(): void {
  console.log('=== Example 1: Validate SharePoint URL ===');

  const urls = [
    'https://contoso.sharepoint.com/sites/marketing',
    'https://www.microsoft.com',
    'https://portal.azure.com',
  ];

  urls.forEach(url => {
    const isValid = isSharePointUrl(url);
    console.log(`${url} -> SharePoint: ${isValid}`);
  });

  console.log('');
}

/**
 * Example 2: Parse SharePoint URL Components
 */
function example2_parseUrl(): void {
  console.log('=== Example 2: Parse URL Components ===');

  const url = 'https://contoso.sharepoint.com/sites/marketing/Shared Documents/Q4 Reports';
  const components = parseSharePointUrl(url);

  if (components) {
    console.log('Original URL:', components.fullUrl);
    console.log('Domain:', components.domain);
    console.log('Site Collection:', components.siteCollectionUrl);
    console.log('Web URL:', components.webUrl);
    console.log('Relative Path:', components.relativePath);
  }

  console.log('');
}

/**
 * Example 3: Detect Context for Different Page Types
 */
function example3_detectPageTypes(): void {
  console.log('=== Example 3: Detect Different Page Types ===');

  const testUrls = [
    {
      label: 'Site Home',
      url: 'https://contoso.sharepoint.com/sites/marketing/SitePages/Home.aspx',
    },
    {
      label: 'Document Library',
      url: 'https://contoso.sharepoint.com/sites/marketing/Shared Documents/Forms/AllItems.aspx',
    },
    {
      label: 'Custom List',
      url: 'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/AllItems.aspx',
    },
    {
      label: 'List Item',
      url: 'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/DispForm.aspx?ID=42',
    },
    {
      label: 'Site Settings',
      url: 'https://contoso.sharepoint.com/sites/marketing/_layouts/15/settings.aspx',
    },
  ];

  testUrls.forEach(({ label, url }) => {
    const result = detectSharePointContext(url);
    if (result.success && result.context) {
      console.log(`${label}:`);
      console.log(`  Page Context: ${result.context.pageContext}`);
      console.log(`  Version: ${result.context.version}`);
    }
  });

  console.log('');
}

/**
 * Example 4: Extract List/Library Information
 */
function example4_extractListInfo(): void {
  console.log('=== Example 4: Extract List Information ===');

  const url = 'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/DispForm.aspx?ID=42&List={12345678-1234-1234-1234-123456789abc}';
  const result = detectSharePointContext(url);

  if (result.success && result.context && result.context.listContext) {
    console.log('List ID:', result.context.listContext.listId);
    console.log('Item ID:', result.context.listContext.itemId);
    console.log('List Name:', result.context.listContext.listName);
  }

  console.log('');
}

/**
 * Example 5: Content Script Usage Pattern
 * This shows how you would use the detector in a browser extension content script
 */
function example5_contentScriptPattern(): void {
  console.log('=== Example 5: Content Script Pattern ===');
  console.log('This example shows the pattern for use in a content script:');
  console.log('');
  console.log('```typescript');
  console.log('// In your content script');
  console.log('import { detectCurrentContext } from "./context";');
  console.log('');
  console.log('// Detect context when page loads');
  console.log('const result = detectCurrentContext(document);');
  console.log('');
  console.log('if (result.success && result.context) {');
  console.log('  const ctx = result.context;');
  console.log('  ');
  console.log('  // Store context for use by extension');
  console.log('  chrome.storage.local.set({ currentContext: ctx });');
  console.log('  ');
  console.log('  // Enable features based on context');
  console.log('  if (ctx.pageContext === PageContext.List || ');
  console.log('      ctx.pageContext === PageContext.DocumentLibrary) {');
  console.log('    // Show quick links for list/library actions');
  console.log('    initializeQuickNav(ctx);');
  console.log('  }');
  console.log('}');
  console.log('```');
  console.log('');
}

/**
 * Example 6: Building Context-Aware URLs
 */
function example6_buildContextAwareUrls(context: SharePointContext): void {
  console.log('=== Example 6: Build Context-Aware URLs ===');

  if (!context.urlComponents) {
    console.log('No URL components available');
    return;
  }

  const { webUrl, siteCollectionUrl } = context.urlComponents;

  // Build various admin/settings URLs
  const urls = {
    siteSettings: `${webUrl}/_layouts/15/settings.aspx`,
    siteContents: `${webUrl}/_layouts/15/viewlsts.aspx`,
    sitePermissions: `${webUrl}/_layouts/15/user.aspx`,
    recycleBin: `${webUrl}/_layouts/15/RecycleBin.aspx`,
    siteFeatures: `${webUrl}/_layouts/15/ManageFeatures.aspx`,
    regionalSettings: `${webUrl}/_layouts/15/RegionalSettings.aspx`,
  };

  console.log('Context-aware URLs for current site:');
  Object.entries(urls).forEach(([name, url]) => {
    console.log(`  ${name}: ${url}`);
  });

  // If we have a list context, build list-specific URLs
  if (context.listContext && context.listContext.listId) {
    const listId = context.listContext.listId;
    const listUrls = {
      listSettings: `${webUrl}/_layouts/15/listedit.aspx?List={${listId}}`,
      listPermissions: `${webUrl}/_layouts/15/user.aspx?List={${listId}}`,
      listVersions: `${webUrl}/_layouts/15/Versions.aspx?List={${listId}}`,
    };

    console.log('\nList-specific URLs:');
    Object.entries(listUrls).forEach(([name, url]) => {
      console.log(`  ${name}: ${url}`);
    });
  }

  console.log('');
}

/**
 * Example 7: Handle Different Environments
 */
function example7_handleEnvironments(): void {
  console.log('=== Example 7: Handle Different Environments ===');

  const environments = [
    {
      name: 'Production Tenant',
      url: 'https://contoso.sharepoint.com/sites/intranet',
    },
    {
      name: 'Development Tenant',
      url: 'https://contosodev.sharepoint.com/sites/intranet',
    },
    {
      name: 'GCC High',
      url: 'https://contoso.sharepoint.us/sites/intranet',
    },
  ];

  environments.forEach(env => {
    const result = detectSharePointContext(env.url);
    if (result.success && result.context) {
      console.log(`${env.name}:`);
      console.log(`  Tenant ID: ${result.context.tenantId}`);
      console.log(`  Domain: ${result.context.urlComponents?.domain}`);
      console.log(`  Site URL: ${result.context.urlComponents?.siteCollectionUrl}`);
    }
  });

  console.log('');
}

/**
 * Run all examples
 */
export function runAllExamples(): void {
  console.log('SharePoint Context Detection - Usage Examples\n');
  console.log('='.repeat(50));
  console.log('');

  example1_validateSharePointUrl();
  example2_parseUrl();
  example3_detectPageTypes();
  example4_extractListInfo();
  example5_contentScriptPattern();

  // Example 6 needs a real context, so we'll create one
  const sampleUrl = 'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/AllItems.aspx';
  const result = detectSharePointContext(sampleUrl);
  if (result.success && result.context) {
    example6_buildContextAwareUrls(result.context);
  }

  example7_handleEnvironments();

  console.log('='.repeat(50));
  console.log('\nAll examples complete!');
}

// Run examples if executed directly
if (require.main === module) {
  runAllExamples();
}
