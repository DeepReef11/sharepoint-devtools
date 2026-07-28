/**
 * SharePoint Detector Tests
 * Validates SharePoint page detection and context extraction
 */

import { describe, it, expect, beforeEach } from '@jest/globals';
import {
  isSharePointPage,
  extractSiteUrl,
  extractWebUrl,
  extractListId,
  detectPageType,
  getSharePointContext,
} from '../src/utils/sharepoint-detector';
import { SharePointPageType, SharePointVersion } from '../src/types/sharepoint';

describe('SharePoint Detector Utilities', () => {
  describe('isSharePointPage', () => {
    it('should detect SharePoint Online URLs', () => {
      const urls = [
        'https://contoso.sharepoint.com/sites/marketing',
        'https://tenant.sharepoint.com',
        'https://tenant-admin.sharepoint.com',
      ];

      urls.forEach((url) => {
        expect(isSharePointPage(url)).toBe(true);
      });
    });

    it('should detect SharePoint GCC URLs', () => {
      const urls = [
        'https://contoso.sharepoint-df.com/sites/marketing',
        'https://tenant.sharepoint-df.com',
      ];

      urls.forEach((url) => {
        expect(isSharePointPage(url)).toBe(true);
      });
    });

    it('should reject non-SharePoint URLs', () => {
      const urls = [
        'https://google.com',
        'https://microsoft.com',
        'https://example.com/sharepoint',
      ];

      urls.forEach((url) => {
        expect(isSharePointPage(url)).toBe(false);
      });
    });
  });

  describe('extractSiteUrl', () => {
    it('should extract site URL from subsite path', () => {
      const url = 'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks';
      const siteUrl = extractSiteUrl(url);
      expect(siteUrl).toBe('https://contoso.sharepoint.com/sites/marketing');
    });

    it('should extract root site URL', () => {
      const url = 'https://contoso.sharepoint.com/Shared%20Documents';
      const siteUrl = extractSiteUrl(url);
      expect(siteUrl).toBe('https://contoso.sharepoint.com');
    });

    it('should handle URLs with query parameters', () => {
      const url = 'https://contoso.sharepoint.com/sites/hr?RootFolder=/Shared';
      const siteUrl = extractSiteUrl(url);
      expect(siteUrl).toBe('https://contoso.sharepoint.com/sites/hr');
    });

    it('should return undefined for non-SharePoint URLs', () => {
      const url = 'https://example.com';
      const siteUrl = extractSiteUrl(url);
      expect(siteUrl).toBeUndefined();
    });
  });

  describe('extractWebUrl', () => {
    it('should extract web URL for root site', () => {
      const url = 'https://contoso.sharepoint.com/sites/marketing';
      const webUrl = extractWebUrl(url);
      expect(webUrl).toBe('https://contoso.sharepoint.com/sites/marketing');
    });

    it('should extract web URL for subweb', () => {
      const url = 'https://contoso.sharepoint.com/sites/marketing/finance/Lists/Budget';
      const webUrl = extractWebUrl(url);
      expect(webUrl).toBe('https://contoso.sharepoint.com/sites/marketing/finance');
    });

    it('should handle root tenant URLs', () => {
      const url = 'https://contoso.sharepoint.com';
      const webUrl = extractWebUrl(url);
      expect(webUrl).toBe('https://contoso.sharepoint.com');
    });
  });

  describe('extractListId', () => {
    it('should extract list ID from List parameter', () => {
      const url =
        'https://contoso.sharepoint.com/_layouts/15/listedit.aspx?List={F3A4B5C6-1234-5678-90AB-CDEF12345678}';
      const listId = extractListId(url);
      expect(listId).toBe('F3A4B5C6-1234-5678-90AB-CDEF12345678');
    });

    it('should extract list ID from ListId parameter', () => {
      const url =
        'https://contoso.sharepoint.com/sites/test?ListId=a1b2c3d4-5678-90ab-cdef-1234567890ab';
      const listId = extractListId(url);
      expect(listId).toBe('A1B2C3D4-5678-90AB-CDEF-1234567890AB');
    });

    it('should extract list name from path', () => {
      const url = 'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/AllItems.aspx';
      const listId = extractListId(url);
      expect(listId).toBe('Tasks');
    });

    it('should return undefined if no list ID found', () => {
      const url = 'https://contoso.sharepoint.com/sites/marketing';
      const listId = extractListId(url);
      expect(listId).toBeUndefined();
    });
  });

  describe('detectPageType', () => {
    it('should detect site settings pages', () => {
      const url = 'https://contoso.sharepoint.com/sites/test/_layouts/15/settings.aspx';
      const pageType = detectPageType(url);
      expect(pageType).toBe(SharePointPageType.SITE_SETTINGS);
    });

    it('should detect list view pages', () => {
      const url = 'https://contoso.sharepoint.com/sites/test/Lists/Tasks/AllItems.aspx?View={guid}';
      const pageType = detectPageType(url);
      expect(pageType).toBe(SharePointPageType.LIST_VIEW);
    });

    it('should detect list form pages', () => {
      const url = 'https://contoso.sharepoint.com/sites/test/Lists/Tasks/Forms/NewForm.aspx';
      const pageType = detectPageType(url);
      expect(pageType).toBe(SharePointPageType.LIST_FORM);
    });

    it('should detect site pages', () => {
      const url = 'https://contoso.sharepoint.com/sites/test/SitePages/Home.aspx';
      const pageType = detectPageType(url);
      expect(pageType).toBe(SharePointPageType.PAGE);
    });

    it('should detect library pages', () => {
      const url = 'https://contoso.sharepoint.com/sites/test/Shared%20Documents';
      const pageType = detectPageType(url);
      expect(pageType).toBe(SharePointPageType.LIBRARY);
    });

    it('should detect admin pages', () => {
      const url = 'https://contoso-admin.sharepoint.com/_layouts/15/settings.aspx';
      const pageType = detectPageType(url);
      expect(pageType).toBe(SharePointPageType.ADMIN);
    });

    it('should detect site home pages', () => {
      const url = 'https://contoso.sharepoint.com/sites/test';
      const pageType = detectPageType(url);
      expect(pageType).toBe(SharePointPageType.SITE_HOME);
    });

    it('should return unknown for unrecognized pages', () => {
      const url = 'https://contoso.sharepoint.com/custom/path';
      const pageType = detectPageType(url);
      expect(pageType).toBe(SharePointPageType.UNKNOWN);
    });
  });

  describe('getSharePointContext', () => {
    it('should return complete context for SharePoint page', () => {
      const url =
        'https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/AllItems.aspx?List={F3A4B5C6-1234-5678-90AB-CDEF12345678}';
      const context = getSharePointContext(url);

      expect(context.isSharePoint).toBe(true);
      expect(context.url).toBe(url);
      expect(context.siteUrl).toBe('https://contoso.sharepoint.com/sites/marketing');
      expect(context.listId).toBe('F3A4B5C6-1234-5678-90AB-CDEF12345678');
      expect(context.pageType).toBe(SharePointPageType.LIST_VIEW);
    });

    it('should return minimal context for non-SharePoint page', () => {
      const url = 'https://example.com';
      const context = getSharePointContext(url);

      expect(context.isSharePoint).toBe(false);
      expect(context.version).toBe(SharePointVersion.UNKNOWN);
      expect(context.url).toBe(url);
      expect(context.siteUrl).toBeUndefined();
      expect(context.webUrl).toBeUndefined();
    });

    it('should handle root SharePoint URLs', () => {
      const url = 'https://contoso.sharepoint.com';
      const context = getSharePointContext(url);

      expect(context.isSharePoint).toBe(true);
      expect(context.siteUrl).toBe('https://contoso.sharepoint.com');
      expect(context.pageType).toBe(SharePointPageType.SITE_HOME);
    });

    it('should handle GCC SharePoint URLs', () => {
      const url = 'https://contoso.sharepoint-df.com/sites/secure';
      const context = getSharePointContext(url);

      expect(context.isSharePoint).toBe(true);
      expect(context.siteUrl).toBe('https://contoso.sharepoint-df.com/sites/secure');
    });
  });
});
