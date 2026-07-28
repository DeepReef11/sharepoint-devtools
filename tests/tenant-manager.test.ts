/**
 * Multi-tenant detection tests.
 *
 * This file previously consisted of exported `console.log` functions that nothing
 * called, plus a single `toBeDefined` assertion — so tenant/cloud detection had no
 * real coverage, which is how the GCC / GCC High domain mapping stayed wrong.
 * These are the same scenarios, expressed as assertions that can actually fail.
 */

import { describe, it, expect } from '@jest/globals';
import { TenantManager } from '../src/managers/tenant-manager';
import { SharePointCloud } from '../src/types/tenant';

const manager = new TenantManager();

describe('TenantManager.detectTenantFromUrl', () => {
  it('detects a commercial tenant', () => {
    const info = manager.detectTenantFromUrl('https://contoso.sharepoint.com/sites/mysite');

    expect(info).not.toBeNull();
    expect(info!.tenantUrl).toBe('https://contoso.sharepoint.com');
    expect(info!.tenantName).toBe('contoso');
    expect(info!.cloud).toBe(SharePointCloud.Commercial);
  });

  it('detects a GCC High tenant on *.sharepoint.us', () => {
    const info = manager.detectTenantFromUrl('https://contoso.sharepoint.us/sites/mysite');

    expect(info!.tenantUrl).toBe('https://contoso.sharepoint.us');
    expect(info!.cloud).toBe(SharePointCloud.GCCHigh);
  });

  it('detects a DoD tenant on *.sharepoint-mil.us', () => {
    const info = manager.detectTenantFromUrl('https://contoso.sharepoint-mil.us/sites/mysite');

    expect(info!.cloud).toBe(SharePointCloud.DoD);
  });

  it('detects a China (21Vianet) tenant', () => {
    const info = manager.detectTenantFromUrl('https://contoso.sharepoint.cn/sites/mysite');

    expect(info!.tenantUrl).toBe('https://contoso.sharepoint.cn');
    expect(info!.cloud).toBe(SharePointCloud.China);
  });

  it('classifies the dogfood ring as Dogfood, not as a government cloud', () => {
    const info = manager.detectTenantFromUrl('https://contoso.sharepoint-df.com/sites/mysite');

    expect(info!.cloud).toBe(SharePointCloud.Dogfood);
  });

  it('treats a GCC tenant as commercial, since GCC shares the worldwide endpoints', () => {
    const info = manager.detectTenantFromUrl('https://agency.sharepoint.com/sites/mysite');

    expect(info!.cloud).toBe(SharePointCloud.Commercial);
  });

  it('normalises the hostname to lower case', () => {
    const info = manager.detectTenantFromUrl('https://CONTOSO.SharePoint.US/sites/mysite');

    expect(info!.hostname).toBe('contoso.sharepoint.us');
    expect(info!.cloud).toBe(SharePointCloud.GCCHigh);
  });

  it('resolves the tenant name from an admin hostname', () => {
    const info = manager.detectTenantFromUrl('https://contoso-admin.sharepoint.com');

    expect(info!.tenantName).toBe('contoso');
  });

  it('returns null for a non-SharePoint URL', () => {
    expect(manager.detectTenantFromUrl('https://example.com/sites/mysite')).toBeNull();
  });

  it('returns null for a lookalike domain', () => {
    expect(
      manager.detectTenantFromUrl('https://contoso.sharepoint.com.attacker.example')
    ).toBeNull();
  });

  it('returns null for malformed input rather than throwing', () => {
    expect(manager.detectTenantFromUrl('not-a-url')).toBeNull();
  });
});
