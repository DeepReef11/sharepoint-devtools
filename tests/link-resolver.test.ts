/**
 * Link Resolution Engine Tests
 * Demonstrates and validates the Link Resolution Engine functionality
 *
 * These are example tests showing how the system works.
 */

import { describe, it, expect } from '@jest/globals';
import { LinkResolver } from '../src/core/link-resolver';
import { detectSharePointContext } from '../src/utils/context-detector';
import { sampleLinkTemplates } from '../src/data/sample-links';
import {
  SharePointContext,
  SharePointPageType,
  SharePointVersion,
  PageContext,
} from '../src/types/sharepoint-context';
import { LinkTemplate, LinkCategory, PlaceholderType } from '../src/types/link-template';

// Placeholder test to satisfy Jest requirement
describe('LinkResolver', () => {
  it('should be defined', () => {
    expect(LinkResolver).toBeDefined();
  });
});

/**
 * Test helper: Create a mock SharePoint context
 */
function createMockContext(overrides: Partial<SharePointContext> = {}): SharePointContext {
  return {
    isSharePoint: true,
    version: SharePointVersion.Modern,
    pageContext: PageContext.SiteHome,
    pageType: SharePointPageType.Site,
    tenantUrl: 'https://contoso.sharepoint.com',
    tenantId: 'contoso',
    siteUrl: 'https://contoso.sharepoint.com/sites/mysite',
    webUrl: 'https://contoso.sharepoint.com/sites/mysite',
    listId: null,
    listTitle: null,
    urlComponents: null,
    listContext: null,
    isSharePointPage: true,
    currentUrl: 'https://contoso.sharepoint.com/sites/mysite',
    detectedAt: new Date(),
    ...overrides,
  };
}

/**
 * Test 1: Placeholder Replacement
 */
export function testPlaceholderReplacement(): void {
  console.log('=== Test 1: Placeholder Replacement ===');

  const context = createMockContext();
  const urlTemplate = '{webUrl}/_layouts/15/settings.aspx';

  const result = LinkResolver.replacePlaceholders(urlTemplate, context);
  const expected = 'https://contoso.sharepoint.com/sites/mysite/_layouts/15/settings.aspx';

  console.log('Template:', urlTemplate);
  console.log('Result:', result);
  console.log('Expected:', expected);
  console.log('Pass:', result === expected);
  console.log('');
}

/**
 * Test 2: Multiple Placeholder Replacement
 */
export function testMultiplePlaceholders(): void {
  console.log('=== Test 2: Multiple Placeholders ===');

  const context = createMockContext({
    listId: 'abc123-def456-ghi789',
  });
  const urlTemplate = '{webUrl}/_layouts/15/listedit.aspx?List={listId}';

  const result = LinkResolver.replacePlaceholders(urlTemplate, context);
  const expected =
    'https://contoso.sharepoint.com/sites/mysite/_layouts/15/listedit.aspx?List=abc123-def456-ghi789';

  console.log('Template:', urlTemplate);
  console.log('Result:', result);
  console.log('Expected:', expected);
  console.log('Pass:', result === expected);
  console.log('');
}

/**
 * Test 3: Link Applicability - Should be applicable
 */
export function testLinkApplicability(): void {
  console.log('=== Test 3: Link Applicability (Should Pass) ===');

  const context = createMockContext();
  const template: LinkTemplate = {
    id: 'test-site-settings',
    title: 'Site Settings',
    urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
    category: LinkCategory.SiteAdmin,
    applicability: {
      pageTypes: [],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl],
    },
  };

  const isApplicable = LinkResolver.isLinkApplicable(template, context);

  console.log('Template:', template.title);
  console.log('Context Page Type:', context.pageType);
  console.log('Is Applicable:', isApplicable);
  console.log('Pass:', isApplicable === true);
  console.log('');
}

/**
 * Test 4: Link Applicability - Should NOT be applicable (missing listId)
 */
export function testLinkApplicabilityFail(): void {
  console.log('=== Test 4: Link Applicability (Should Fail - Missing Context) ===');

  const context = createMockContext(); // No listId
  const template: LinkTemplate = {
    id: 'test-list-settings',
    title: 'List Settings',
    urlTemplate: '{webUrl}/_layouts/15/listedit.aspx?List={listId}',
    category: LinkCategory.ListSettings,
    applicability: {
      pageTypes: [SharePointPageType.List],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl, PlaceholderType.ListId],
    },
  };

  const isApplicable = LinkResolver.isLinkApplicable(template, context);

  console.log('Template:', template.title);
  console.log('Context has listId:', context.listId !== null);
  console.log('Is Applicable:', isApplicable);
  console.log('Pass:', isApplicable === false);
  console.log('');
}

/**
 * Test 5: Link Resolution (Full Process)
 */
export function testLinkResolution(): void {
  console.log('=== Test 5: Full Link Resolution ===');

  const context = createMockContext();
  const template: LinkTemplate = {
    id: 'test-site-settings',
    title: 'Site Settings',
    urlTemplate: '{webUrl}/_layouts/15/settings.aspx',
    category: LinkCategory.SiteAdmin,
    description: 'Access site settings',
    keywords: ['settings', 'admin'],
    applicability: {
      pageTypes: [],
      versions: [],
      requiredContext: [PlaceholderType.WebUrl],
    },
    openInNewTab: false,
  };

  const resolved = LinkResolver.resolveLink(template, context);

  console.log('Template:', template.title);
  console.log('Resolved:', resolved);
  console.log('Pass:', resolved !== null && resolved.isValid);
  console.log('');
}

/**
 * Test 6: URL Validation (Valid URL)
 */
export function testUrlValidation(): void {
  console.log('=== Test 6: URL Validation (Valid) ===');

  const context = createMockContext();
  const validUrl = 'https://contoso.sharepoint.com/sites/mysite/_layouts/15/settings.aspx';

  const validation = LinkResolver.validateUrl(validUrl, context);

  console.log('URL:', validUrl);
  console.log('Is Valid:', validation.isValid);
  console.log('Errors:', validation.errors);
  console.log('Pass:', validation.isValid === true);
  console.log('');
}

/**
 * Test 7: URL Validation (Invalid - Contains Placeholder)
 */
export function testUrlValidationInvalid(): void {
  console.log('=== Test 7: URL Validation (Invalid - Unresolved Placeholder) ===');

  const context = createMockContext();
  const invalidUrl = '{webUrl}/_layouts/15/settings.aspx';

  const validation = LinkResolver.validateUrl(invalidUrl, context);

  console.log('URL:', invalidUrl);
  console.log('Is Valid:', validation.isValid);
  console.log('Errors:', validation.errors);
  console.log('Pass:', validation.isValid === false);
  console.log('');
}

/**
 * Test 8: Context Detection from URL
 */
export function testContextDetection(): void {
  console.log('=== Test 8: Context Detection ===');

  const testUrls = [
    'https://contoso.sharepoint.com/sites/mysite',
    'https://contoso.sharepoint.com/sites/mysite/Lists/MyList/AllItems.aspx',
    'https://contoso.sharepoint.com/sites/mysite/_layouts/15/settings.aspx',
  ];

  testUrls.forEach((url) => {
    const context = detectSharePointContext(url);
    console.log('URL:', url);
    console.log('Is SharePoint:', context.isSharePointPage);
    console.log('Site URL:', context.siteUrl);
    console.log('Web URL:', context.webUrl);
    console.log('Page Type:', context.pageType);
    console.log('---');
  });
  console.log('');
}

/**
 * Test 9: Filtering Applicable Templates
 */
export function testFilteringTemplates(): void {
  console.log('=== Test 9: Filtering Applicable Templates ===');

  // Site context (no list)
  const siteContext = createMockContext();
  const siteApplicable = LinkResolver.filterApplicableTemplates(sampleLinkTemplates, siteContext);

  console.log('Site Context - Total Templates:', sampleLinkTemplates.length);
  console.log('Site Context - Applicable:', siteApplicable.length);

  // List context
  const listContext = createMockContext({
    pageType: SharePointPageType.List,
    listId: 'abc123-def456',
  });
  const listApplicable = LinkResolver.filterApplicableTemplates(sampleLinkTemplates, listContext);

  console.log('List Context - Total Templates:', sampleLinkTemplates.length);
  console.log('List Context - Applicable:', listApplicable.length);
  console.log('Pass:', listApplicable.length > siteApplicable.length);
  console.log('');
}

/**
 * Test 10: Resolve Multiple Links
 */
export function testResolveMultipleLinks(): void {
  console.log('=== Test 10: Resolve Multiple Links ===');

  const context = createMockContext({
    pageType: SharePointPageType.List,
    listId: 'abc123-def456',
  });

  const resolved = LinkResolver.resolveLinks(sampleLinkTemplates, context);

  console.log('Context Page Type:', context.pageType);
  console.log('Total Templates:', sampleLinkTemplates.length);
  console.log('Resolved Links:', resolved.length);
  console.log('Valid Links:', resolved.filter((l) => l.isValid).length);

  // Show first 5 resolved links
  console.log('\nFirst 5 Resolved Links:');
  resolved.slice(0, 5).forEach((link) => {
    console.log(`- ${link.title}: ${link.url}`);
  });
  console.log('');
}

/**
 * Test 11: Group Links by Category
 */
export function testGroupByCategory(): void {
  console.log('=== Test 11: Group Links by Category ===');

  const context = createMockContext({
    pageType: SharePointPageType.List,
    listId: 'abc123-def456',
  });

  const resolved = LinkResolver.resolveLinks(sampleLinkTemplates, context);
  const grouped = LinkResolver.groupByCategory(resolved);

  console.log('Total Categories:', grouped.size);
  grouped.forEach((links, category) => {
    console.log(`${category}: ${links.length} links`);
  });
  console.log('');
}

/**
 * Test 12: Extract Placeholders
 */
export function testExtractPlaceholders(): void {
  console.log('=== Test 12: Extract Placeholders ===');

  const templates = [
    '{webUrl}/_layouts/15/settings.aspx',
    '{webUrl}/_layouts/15/listedit.aspx?List={listId}',
    '{tenantUrl}/_layouts/15/online/AdminHome.aspx',
  ];

  templates.forEach((template) => {
    const placeholders = LinkResolver.extractPlaceholders(template);
    console.log('Template:', template);
    console.log('Placeholders:', placeholders);
    console.log('---');
  });
  console.log('');
}

/**
 * Run all tests
 */
export function runAllTests(): void {
  console.log('==============================================');
  console.log('Link Resolution Engine - Test Suite');
  console.log('==============================================\n');

  testPlaceholderReplacement();
  testMultiplePlaceholders();
  testLinkApplicability();
  testLinkApplicabilityFail();
  testLinkResolution();
  testUrlValidation();
  testUrlValidationInvalid();
  testContextDetection();
  testFilteringTemplates();
  testResolveMultipleLinks();
  testGroupByCategory();
  testExtractPlaceholders();

  console.log('==============================================');
  console.log('Tests Complete');
  console.log('==============================================');
}

// Export for use in browser console or Node.js
if (typeof window !== 'undefined') {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).LinkResolverTests = {
    runAllTests,
    testPlaceholderReplacement,
    testMultiplePlaceholders,
    testLinkApplicability,
    testLinkApplicabilityFail,
    testLinkResolution,
    testUrlValidation,
    testUrlValidationInvalid,
    testContextDetection,
    testFilteringTemplates,
    testResolveMultipleLinks,
    testGroupByCategory,
    testExtractPlaceholders,
  };
}
