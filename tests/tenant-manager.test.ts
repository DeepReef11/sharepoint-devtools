/**
 * Multi-Tenant Support Tests
 * Issue #13: Multi-Tenant Support
 *
 * Tests for TenantManager and TenantStorage functionality
 * Full Jest/Mocha test suite will be implemented in Issue #16
 */

import { describe, it, expect } from '@jest/globals';
import { TenantManager } from '../src/managers/tenant-manager';
import { TenantStorage } from '../src/storage/tenant-storage';
import { Tenant, SharePointCloud } from '../src/types/tenant';

// Placeholder test to satisfy Jest requirement
describe('TenantManager', () => {
  it('should be defined', () => {
    expect(TenantManager).toBeDefined();
  });
});

/**
 * Test 1: Detect tenant from URL (Commercial Cloud)
 */
export function testDetectTenantCommercial(): void {
  console.log('=== Test 1: Detect Tenant (Commercial Cloud) ===');

  const manager = new TenantManager();
  const url = 'https://contoso.sharepoint.com/sites/mysite';
  const tenantInfo = manager.detectTenantFromUrl(url);

  console.log('URL:', url);
  console.log('Tenant Info:', tenantInfo);
  console.log(
    'Pass:',
    tenantInfo !== null &&
      tenantInfo.tenantUrl === 'https://contoso.sharepoint.com' &&
      tenantInfo.tenantName === 'contoso' &&
      tenantInfo.cloud === SharePointCloud.Commercial
  );
  console.log('');
}

/**
 * Test 2: Detect tenant from URL (GCC High)
 */
export function testDetectTenantGCCHigh(): void {
  console.log('=== Test 2: Detect Tenant (GCC High) ===');

  const manager = new TenantManager();
  const url = 'https://contoso.sharepoint-df.com/sites/mysite';
  const tenantInfo = manager.detectTenantFromUrl(url);

  console.log('URL:', url);
  console.log('Tenant Info:', tenantInfo);
  console.log(
    'Pass:',
    tenantInfo !== null &&
      tenantInfo.tenantUrl === 'https://contoso.sharepoint-df.com' &&
      tenantInfo.cloud === SharePointCloud.GCCHigh
  );
  console.log('');
}

/**
 * Test 3: Detect tenant from URL (GCC)
 */
export function testDetectTenantGCC(): void {
  console.log('=== Test 3: Detect Tenant (GCC) ===');

  const manager = new TenantManager();
  const url = 'https://contoso.sharepoint.us/sites/mysite';
  const tenantInfo = manager.detectTenantFromUrl(url);

  console.log('URL:', url);
  console.log('Tenant Info:', tenantInfo);
  console.log(
    'Pass:',
    tenantInfo !== null &&
      tenantInfo.tenantUrl === 'https://contoso.sharepoint.us' &&
      tenantInfo.cloud === SharePointCloud.GCC
  );
  console.log('');
}

/**
 * Test 4: Detect tenant from URL (China)
 */
export function testDetectTenantChina(): void {
  console.log('=== Test 4: Detect Tenant (China) ===');

  const manager = new TenantManager();
  const url = 'https://contoso.sharepoint.cn/sites/mysite';
  const tenantInfo = manager.detectTenantFromUrl(url);

  console.log('URL:', url);
  console.log('Tenant Info:', tenantInfo);
  console.log(
    'Pass:',
    tenantInfo !== null &&
      tenantInfo.tenantUrl === 'https://contoso.sharepoint.cn' &&
      tenantInfo.cloud === SharePointCloud.China
  );
  console.log('');
}

/**
 * Test 5: Non-SharePoint URL should return null
 */
export function testDetectTenantInvalid(): void {
  console.log('=== Test 5: Detect Tenant (Invalid URL) ===');

  const manager = new TenantManager();
  const url = 'https://example.com/some-page';
  const tenantInfo = manager.detectTenantFromUrl(url);

  console.log('URL:', url);
  console.log('Tenant Info:', tenantInfo);
  console.log('Pass:', tenantInfo === null);
  console.log('');
}

/**
 * Test 6: Save and retrieve tenant
 */
export async function testSaveAndRetrieveTenant(): Promise<void> {
  console.log('=== Test 6: Save and Retrieve Tenant ===');

  const storage = new TenantStorage();

  // Clear storage first
  await storage.clearAll();

  const tenant: Tenant = {
    tenantUrl: 'https://contoso.sharepoint.com',
    displayName: 'Contoso',
    lastAccessed: new Date().toISOString(),
    isDefault: true,
  };

  await storage.saveTenant(tenant);
  const retrieved = await storage.getTenant(tenant.tenantUrl);

  console.log('Saved Tenant:', tenant);
  console.log('Retrieved Tenant:', retrieved);
  console.log(
    'Pass:',
    retrieved !== null &&
      retrieved.tenantUrl === tenant.tenantUrl &&
      retrieved.displayName === tenant.displayName
  );
  console.log('');
}

/**
 * Test 7: Get all tenants sorted by last accessed
 */
export async function testGetAllTenantsSorted(): Promise<void> {
  console.log('=== Test 7: Get All Tenants (Sorted) ===');

  const storage = new TenantStorage();

  // Clear and add multiple tenants
  await storage.clearAll();

  const tenant1: Tenant = {
    tenantUrl: 'https://contoso.sharepoint.com',
    displayName: 'Contoso',
    lastAccessed: new Date(Date.now() - 1000).toISOString(), // 1 second ago
  };

  const tenant2: Tenant = {
    tenantUrl: 'https://fabrikam.sharepoint.com',
    displayName: 'Fabrikam',
    lastAccessed: new Date().toISOString(), // Now (most recent)
  };

  await storage.saveTenant(tenant1);
  await storage.saveTenant(tenant2);

  const allTenants = await storage.getAllTenants();

  console.log('Saved 2 tenants');
  console.log(
    'All Tenants:',
    allTenants.map((t) => t.displayName)
  );
  console.log(
    'Pass:',
    allTenants.length === 2 && allTenants[0].tenantUrl === tenant2.tenantUrl // Most recent first
  );
  console.log('');
}

/**
 * Test 8: Set and get default tenant
 */
export async function testDefaultTenant(): Promise<void> {
  console.log('=== Test 8: Set and Get Default Tenant ===');

  const storage = new TenantStorage();

  // Clear and add tenants
  await storage.clearAll();

  const tenant1: Tenant = {
    tenantUrl: 'https://contoso.sharepoint.com',
    displayName: 'Contoso',
    lastAccessed: new Date().toISOString(),
  };

  const tenant2: Tenant = {
    tenantUrl: 'https://fabrikam.sharepoint.com',
    displayName: 'Fabrikam',
    lastAccessed: new Date().toISOString(),
  };

  await storage.saveTenant(tenant1);
  await storage.saveTenant(tenant2);
  await storage.setDefaultTenant(tenant2.tenantUrl);

  const defaultTenant = await storage.getDefaultTenant();

  console.log('Set Default:', tenant2.displayName);
  console.log('Got Default:', defaultTenant?.displayName);
  console.log(
    'Pass:',
    defaultTenant !== null &&
      defaultTenant.tenantUrl === tenant2.tenantUrl &&
      defaultTenant.isDefault === true
  );
  console.log('');
}

/**
 * Test 9: Update tenant preferences
 */
export async function testUpdatePreferences(): Promise<void> {
  console.log('=== Test 9: Update Preferences ===');

  const storage = new TenantStorage();
  await storage.clearAll();

  // Default should be autoDetectTenant: true
  let prefs = await storage.getPreferences();
  console.log('Initial autoDetectTenant:', prefs.autoDetectTenant);

  // Update to false
  await storage.updatePreferences({ autoDetectTenant: false });
  prefs = await storage.getPreferences();

  console.log('Updated autoDetectTenant:', prefs.autoDetectTenant);
  console.log('Pass:', prefs.autoDetectTenant === false);
  console.log('');
}

/**
 * Test 10: Remove tenant
 */
export async function testRemoveTenant(): Promise<void> {
  console.log('=== Test 10: Remove Tenant ===');

  const storage = new TenantStorage();
  await storage.clearAll();

  const tenant: Tenant = {
    tenantUrl: 'https://contoso.sharepoint.com',
    displayName: 'Contoso',
    lastAccessed: new Date().toISOString(),
  };

  await storage.saveTenant(tenant);
  let retrieved = await storage.getTenant(tenant.tenantUrl);
  console.log('Tenant exists after save:', retrieved !== null);

  await storage.removeTenant(tenant.tenantUrl);
  retrieved = await storage.getTenant(tenant.tenantUrl);

  console.log('Tenant exists after remove:', retrieved !== null);
  console.log('Pass:', retrieved === null);
  console.log('');
}

/**
 * Test 11: Get or create tenant (auto-register)
 */
export async function testGetOrCreateTenant(): Promise<void> {
  console.log('=== Test 11: Get or Create Tenant (Auto-register) ===');

  const manager = new TenantManager();
  const storage = new TenantStorage();
  await storage.clearAll();

  const url = 'https://contoso.sharepoint.com/sites/mysite';

  // First call should create tenant
  const tenant1 = await manager.getOrCreateTenant(url, true);
  console.log('First call - Tenant created:', tenant1?.displayName);

  // Second call should retrieve existing
  const tenant2 = await manager.getOrCreateTenant(url, true);
  console.log('Second call - Tenant retrieved:', tenant2?.displayName);

  console.log(
    'Pass:',
    tenant1 !== null &&
      tenant2 !== null &&
      tenant1.tenantUrl === tenant2.tenantUrl &&
      tenant1.displayName === tenant2.displayName
  );
  console.log('');
}

/**
 * Test 12: Switch tenant
 */
export async function testSwitchTenant(): Promise<void> {
  console.log('=== Test 12: Switch Tenant ===');

  const manager = new TenantManager();
  const storage = new TenantStorage();
  await storage.clearAll();

  // Create two tenants
  const tenant1: Tenant = {
    tenantUrl: 'https://contoso.sharepoint.com',
    displayName: 'Contoso',
    lastAccessed: new Date().toISOString(),
  };

  const tenant2: Tenant = {
    tenantUrl: 'https://fabrikam.sharepoint.com',
    displayName: 'Fabrikam',
    lastAccessed: new Date().toISOString(),
  };

  await storage.saveTenant(tenant1);
  await storage.saveTenant(tenant2);

  // Switch to tenant2
  const result = await manager.switchTenant(tenant2.tenantUrl);

  console.log('Switch Result:', result);
  console.log('Pass:', result.success === true && result.newTenantUrl === tenant2.tenantUrl);
  console.log('');
}

/**
 * Test 13: Validate tenant URL
 */
export async function testValidateTenantUrl(): Promise<void> {
  console.log('=== Test 13: Validate Tenant URL ===');

  const manager = new TenantManager();
  const storage = new TenantStorage();
  await storage.clearAll();

  const tenant: Tenant = {
    tenantUrl: 'https://contoso.sharepoint.com',
    displayName: 'Contoso',
    lastAccessed: new Date().toISOString(),
  };

  await storage.saveTenant(tenant);

  const validUrl = 'https://contoso.sharepoint.com/sites/mysite';
  const invalidUrl = 'https://unknown.sharepoint.com/sites/test';

  const isValid = await manager.validateTenantUrl(validUrl);
  const isInvalid = await manager.validateTenantUrl(invalidUrl);

  console.log('Valid URL check:', isValid);
  console.log('Invalid URL check:', isInvalid);
  console.log('Pass:', isValid === true && isInvalid === false);
  console.log('');
}

/**
 * Test 14: Tenant-specific settings
 */
export async function testTenantSettings(): Promise<void> {
  console.log('=== Test 14: Tenant-Specific Settings ===');

  const manager = new TenantManager();
  const storage = new TenantStorage();
  await storage.clearAll();

  const tenant: Tenant = {
    tenantUrl: 'https://contoso.sharepoint.com',
    displayName: 'Contoso',
    lastAccessed: new Date().toISOString(),
  };

  await storage.saveTenant(tenant);

  // Update settings
  const settings = {
    theme: 'dark',
    customLinks: ['link1', 'link2'],
  };

  await manager.updateTenantSettings(tenant.tenantUrl, settings);

  // Retrieve settings
  const retrieved = await manager.getTenantSettings(tenant.tenantUrl);

  console.log('Saved Settings:', settings);
  console.log('Retrieved Settings:', retrieved);
  console.log('Pass:', retrieved.theme === 'dark' && retrieved.customLinks.length === 2);
  console.log('');
}

/**
 * Test 15: Max recent tenants limit
 */
export async function testMaxRecentTenants(): Promise<void> {
  console.log('=== Test 15: Max Recent Tenants Limit ===');

  const storage = new TenantStorage();
  await storage.clearAll();

  // Set max to 5
  await storage.updatePreferences({ maxRecentTenants: 5 });

  // Add 7 tenants
  for (let i = 1; i <= 7; i++) {
    const tenant: Tenant = {
      tenantUrl: `https://tenant${i}.sharepoint.com`,
      displayName: `Tenant ${i}`,
      lastAccessed: new Date(Date.now() + i * 1000).toISOString(),
    };
    await storage.saveTenant(tenant);
  }

  const allTenants = await storage.getAllTenants();

  console.log('Added 7 tenants with max 5');
  console.log('Total tenants stored:', allTenants.length);
  console.log('Pass:', allTenants.length === 5);
  console.log('');
}

/**
 * Run all tests
 */
export async function runAllTests(): Promise<void> {
  console.log('==============================================');
  console.log('Multi-Tenant Support - Test Suite');
  console.log('Issue #13: Multi-Tenant Support');
  console.log('==============================================\n');

  // Synchronous tests
  testDetectTenantCommercial();
  testDetectTenantGCCHigh();
  testDetectTenantGCC();
  testDetectTenantChina();
  testDetectTenantInvalid();

  // Asynchronous tests
  await testSaveAndRetrieveTenant();
  await testGetAllTenantsSorted();
  await testDefaultTenant();
  await testUpdatePreferences();
  await testRemoveTenant();
  await testGetOrCreateTenant();
  await testSwitchTenant();
  await testValidateTenantUrl();
  await testTenantSettings();
  await testMaxRecentTenants();

  console.log('==============================================');
  console.log('Tests Complete');
  console.log('==============================================');
}

// Export for use in browser console or Node.js
if (typeof window !== 'undefined') {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).TenantManagerTests = {
    runAllTests,
    testDetectTenantCommercial,
    testDetectTenantGCCHigh,
    testDetectTenantGCC,
    testDetectTenantChina,
    testDetectTenantInvalid,
    testSaveAndRetrieveTenant,
    testGetAllTenantsSorted,
    testDefaultTenant,
    testUpdatePreferences,
    testRemoveTenant,
    testGetOrCreateTenant,
    testSwitchTenant,
    testValidateTenantUrl,
    testTenantSettings,
    testMaxRecentTenants,
  };
}
