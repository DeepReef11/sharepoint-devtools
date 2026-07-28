/**
 * Cloud detection tests.
 *
 * Domains are taken from Microsoft's published endpoint lists:
 *   GCC High        *.sharepoint.us       (Microsoft 365 U.S. Government GCC High endpoints)
 *   DoD             *.sharepoint-mil.us   (Microsoft 365 U.S. Government DoD endpoints)
 *   GCC             *.sharepoint.com      (served by the Worldwide endpoints)
 *   China/21Vianet  *.sharepoint.cn
 *   Dogfood         *.sharepoint-df.com   (Microsoft-internal first-release ring)
 */

import { describe, it, expect } from '@jest/globals';
import {
  isSharePointHostname,
  isSharePointUrl,
  detectCloud,
  extractTenantName,
  getTenantAdminUrl,
  SHAREPOINT_HOST_PATTERNS,
} from '../src/utils/sharepoint-clouds';
import { SharePointCloud } from '../src/types/tenant';

describe('isSharePointHostname', () => {
  it.each([
    'contoso.sharepoint.com',
    'contoso.sharepoint-df.com',
    'contoso.sharepoint.us',
    'contoso.sharepoint-mil.us',
    'contoso.sharepoint.cn',
    'contoso-admin.sharepoint.com',
    'contoso-admin.sharepoint.us',
  ])('accepts %s', (hostname) => {
    expect(isSharePointHostname(hostname)).toBe(true);
  });

  it.each([
    'example.com',
    'sharepoint.com.evil.example',
    'notsharepoint.us',
    'contoso.onmicrosoft.com',
    '',
  ])('rejects %s', (hostname) => {
    expect(isSharePointHostname(hostname)).toBe(false);
  });

  it('is case-insensitive', () => {
    expect(isSharePointHostname('CONTOSO.SharePoint.US')).toBe(true);
  });

  it('does not match a lookalike domain that merely contains the suffix', () => {
    // A substring check (the previous implementation) would wrongly accept this.
    expect(isSharePointHostname('contoso.sharepoint.com.attacker.example')).toBe(false);
  });
});

describe('detectCloud', () => {
  it('maps GCC High to *.sharepoint.us', () => {
    expect(detectCloud('contoso.sharepoint.us')).toBe(SharePointCloud.GCCHigh);
  });

  it('maps DoD to *.sharepoint-mil.us', () => {
    // Must win over the bare .sharepoint.us suffix.
    expect(detectCloud('contoso.sharepoint-mil.us')).toBe(SharePointCloud.DoD);
  });

  it('maps China to *.sharepoint.cn', () => {
    expect(detectCloud('contoso.sharepoint.cn')).toBe(SharePointCloud.China);
  });

  it('maps the dogfood ring to *.sharepoint-df.com, not to a government cloud', () => {
    expect(detectCloud('contoso.sharepoint-df.com')).toBe(SharePointCloud.Dogfood);
  });

  it('maps commercial to *.sharepoint.com', () => {
    expect(detectCloud('contoso.sharepoint.com')).toBe(SharePointCloud.Commercial);
  });

  it('reports GCC tenants as Commercial, since they share the worldwide endpoints', () => {
    // GCC is not distinguishable from commercial by hostname; asserting this
    // keeps anyone from "fixing" it back to a bogus GCC-specific domain.
    expect(detectCloud('gcctenant.sharepoint.com')).toBe(SharePointCloud.Commercial);
  });

  it('returns Unknown for non-SharePoint hosts', () => {
    expect(detectCloud('example.com')).toBe(SharePointCloud.Unknown);
  });
});

describe('extractTenantName', () => {
  it.each([
    ['contoso.sharepoint.com', 'contoso'],
    ['contoso.sharepoint.us', 'contoso'],
    ['contoso-admin.sharepoint.com', 'contoso'],
    ['contoso-admin.sharepoint.us', 'contoso'],
  ])('extracts %s -> %s', (hostname, expected) => {
    expect(extractTenantName(hostname)).toBe(expected);
  });
});

describe('getTenantAdminUrl', () => {
  it.each([
    ['contoso.sharepoint.com', 'https://contoso-admin.sharepoint.com'],
    ['contoso.sharepoint.us', 'https://contoso-admin.sharepoint.us'],
    ['contoso.sharepoint-mil.us', 'https://contoso-admin.sharepoint-mil.us'],
    ['contoso.sharepoint.cn', 'https://contoso-admin.sharepoint.cn'],
    ['contoso.sharepoint-df.com', 'https://contoso-admin.sharepoint-df.com'],
  ])('builds the admin URL for %s', (hostname, expected) => {
    expect(getTenantAdminUrl(hostname)).toBe(expected);
  });

  it('is idempotent when given an admin hostname', () => {
    expect(getTenantAdminUrl('contoso-admin.sharepoint.us')).toBe(
      'https://contoso-admin.sharepoint.us'
    );
  });

  it('returns an empty string for non-SharePoint hosts', () => {
    expect(getTenantAdminUrl('example.com')).toBe('');
  });
});

describe('isSharePointUrl', () => {
  it('accepts a full URL on any cloud', () => {
    expect(isSharePointUrl('https://contoso.sharepoint.us/sites/team/Lists/Tasks')).toBe(true);
  });

  it('rejects malformed input without throwing', () => {
    expect(isSharePointUrl('not a url')).toBe(false);
  });
});

describe('manifest coverage', () => {
  // A cloud detected in code but absent from the manifest is unreachable: the
  // content script is never injected, so the detection can never run.
  it('declares a host pattern for every cloud the code detects', () => {
    const detectable = [
      'contoso.sharepoint.com',
      'contoso.sharepoint-df.com',
      'contoso.sharepoint.us',
      'contoso.sharepoint-mil.us',
      'contoso.sharepoint.cn',
    ];

    for (const hostname of detectable) {
      const suffix = hostname.replace('contoso', '');
      const covered = SHAREPOINT_HOST_PATTERNS.some((p) => p.includes(suffix + '/*'));
      expect(covered).toBe(true);
    }
  });

  it('matches the patterns declared in public/manifest.json', () => {
    // Guards against the manifest and the code drifting apart again.
    const manifest = require('../public/manifest.json');
    expect(manifest.host_permissions.sort()).toEqual([...SHAREPOINT_HOST_PATTERNS].sort());
    expect(manifest.content_scripts[0].matches.sort()).toEqual(
      [...SHAREPOINT_HOST_PATTERNS].sort()
    );
  });
});
