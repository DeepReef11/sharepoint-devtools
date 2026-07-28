# Multi-Tenant Support

**Multi-Tenant Support**
**Status**: ✅ Implemented

## Overview

The SharePoint DevTools extension now supports multiple SharePoint tenants with automatic detection and tenant-specific settings. This feature enables users who work across multiple SharePoint organizations to seamlessly switch between tenants while maintaining separate configurations for each.

## Features

### 1. Automatic Tenant Detection

The extension automatically detects the SharePoint tenant from the current URL and registers it for future use.

**Supported SharePoint Clouds:**
- **Commercial**: `*.sharepoint.com`
- **GCC High**: `*.sharepoint-df.com`
- **GCC**: `*.sharepoint.us`
- **China**: `*.sharepoint.cn`

### 2. Tenant Registration

When you visit a SharePoint site, the extension:
1. Extracts the tenant URL (e.g., `https://contoso.sharepoint.com`)
2. Identifies the tenant name and cloud environment
3. Automatically registers the tenant with a display name
4. Tracks when the tenant was last accessed

### 3. Tenant Storage

Tenant information is stored persistently using `chrome.storage.sync`, which:
- Syncs across devices where you're signed into Chrome/Edge
- Maintains a list of up to 10 recent tenants (configurable)
- Stores tenant-specific settings and preferences
- Preserves default tenant selection

### 4. Tenant Switching

Users can switch between registered tenants:
- View all known tenants sorted by last accessed
- Set a default tenant for the extension
- Manually switch to a different tenant
- Remove tenants from the registry

### 5. Tenant-Specific Settings

Each tenant can have its own:
- Custom display name
- Tenant-specific custom links
- Individual preferences and settings
- Isolated configuration

## Architecture

### Core Components

#### 1. Type Definitions (`src/types/tenant.ts`)

Defines TypeScript interfaces for tenant data:

```typescript
interface Tenant {
  tenantUrl: string;
  displayName?: string;
  lastAccessed: string;
  isDefault?: boolean;
  customLinks?: any[];
  settings?: Record<string, any>;
}

interface TenantPreferences {
  autoDetectTenant: boolean;
  defaultTenantUrl?: string;
  maxRecentTenants?: number;
}

enum SharePointCloud {
  Commercial = 'commercial',
  GCCHigh = 'gcc-high',
  GCC = 'gcc',
  China = 'china',
  Unknown = 'unknown'
}
```

#### 2. Tenant Storage (`src/storage/tenant-storage.ts`)

Manages persistence using Chrome Storage API:

**Key Methods:**
- `loadData()`: Loads all tenant data from storage
- `saveData()`: Saves tenant data to storage
- `getTenant(tenantUrl)`: Retrieves a specific tenant
- `saveTenant(tenant)`: Adds or updates a tenant
- `removeTenant(tenantUrl)`: Removes a tenant
- `getAllTenants()`: Gets all tenants sorted by last accessed
- `getDefaultTenant()`: Gets the default tenant
- `setDefaultTenant(tenantUrl)`: Sets a tenant as default
- `getPreferences()`: Gets user preferences
- `updatePreferences()`: Updates preferences
- `touchTenant(tenantUrl)`: Updates last accessed time

**Storage Structure:**
```json
{
  "sharepoint_quicknav_tenants": {
    "tenants": [
      {
        "tenantUrl": "https://contoso.sharepoint.com",
        "displayName": "Contoso",
        "lastAccessed": "2026-01-15T09:00:00.000Z",
        "isDefault": true,
        "settings": {}
      }
    ],
    "preferences": {
      "autoDetectTenant": true,
      "maxRecentTenants": 10
    },
    "version": 1
  }
}
```

#### 3. Tenant Manager (`src/managers/tenant-manager.ts`)

Coordinates tenant detection, switching, and management:

**Key Methods:**
- `detectTenantFromUrl(url)`: Extracts tenant info from URL
- `getOrCreateTenant(url, autoRegister)`: Gets or creates tenant
- `getCurrentTenant()`: Gets the active tenant
- `switchTenant(tenantUrl)`: Switches to a different tenant
- `getAllTenants()`: Lists all known tenants
- `updateTenant(tenantUrl, updates)`: Updates tenant info
- `removeTenant(tenantUrl)`: Removes a tenant
- `setDefaultTenant(tenantUrl)`: Sets default tenant
- `isAutoDetectEnabled()`: Checks auto-detect setting
- `validateTenantUrl(url)`: Validates if URL belongs to known tenant
- `getTenantSettings(tenantUrl)`: Gets tenant-specific settings
- `updateTenantSettings(tenantUrl, settings)`: Updates settings

#### 4. Context Detector Integration (`src/utils/context-detector.ts`)

Enhanced context detection with tenant awareness:

**New Functions:**
- `detectContextWithTenant(url)`: Detects context with tenant data
- `getCurrentContextWithTenant()`: Gets current context with tenant
- `isKnownTenant(context)`: Validates if tenant is registered

## Usage Examples

### Detect and Register Tenant

```typescript
import { tenantManager } from './managers/tenant-manager';

// Auto-detect and register tenant from URL
const url = 'https://contoso.sharepoint.com/sites/mysite';
const tenant = await tenantManager.getOrCreateTenant(url, true);

console.log(tenant.displayName); // "contoso"
console.log(tenant.tenantUrl);   // "https://contoso.sharepoint.com"
```

### Get Context with Tenant

```typescript
import { detectContextWithTenant } from './utils/context-detector';

// Detect SharePoint context with tenant info
const { context, tenant } = await detectContextWithTenant();

if (tenant) {
  console.log(`On tenant: ${tenant.displayName}`);
  console.log(`Site URL: ${context.siteUrl}`);
}
```

### Switch Tenants

```typescript
import { tenantManager } from './managers/tenant-manager';

// Get all known tenants
const tenants = await tenantManager.getAllTenants();
console.log('Known tenants:', tenants.map(t => t.displayName));

// Switch to a different tenant
const result = await tenantManager.switchTenant('https://fabrikam.sharepoint.com');
if (result.success) {
  console.log('Switched to:', result.newTenantUrl);
}
```

### Manage Tenant Settings

```typescript
import { tenantManager } from './managers/tenant-manager';

// Get tenant-specific settings
const settings = await tenantManager.getTenantSettings(
  'https://contoso.sharepoint.com'
);

// Update settings
await tenantManager.updateTenantSettings(
  'https://contoso.sharepoint.com',
  {
    theme: 'dark',
    customLinks: [...]
  }
);
```

### Configure Preferences

```typescript
import { tenantStorage } from './storage/tenant-storage';

// Disable auto-detection
await tenantStorage.updatePreferences({
  autoDetectTenant: false
});

// Set max recent tenants
await tenantStorage.updatePreferences({
  maxRecentTenants: 20
});

// Set default tenant
await tenantStorage.setDefaultTenant('https://contoso.sharepoint.com');
```

## Testing

Comprehensive tests are available in `src/tests/tenant-manager.test.ts`:

**Test Coverage:**
- Tenant detection for all cloud environments
- Save and retrieve tenants
- Default tenant management
- Tenant switching
- Preferences management
- Tenant removal
- Auto-registration
- Tenant validation
- Tenant-specific settings
- Max recent tenants limit

**Run Tests:**
```javascript
// In browser console after loading extension
TenantManagerTests.runAllTests();
```

## User Preferences

### Auto-Detection (Default: Enabled)

When enabled, the extension automatically registers new tenants as you visit them.

```typescript
// Enable auto-detection
await tenantManager.setAutoDetect(true);

// Disable auto-detection
await tenantManager.setAutoDetect(false);
```

### Max Recent Tenants (Default: 10)

Limits the number of tenants stored. When exceeded, the least recently accessed tenants are removed.

```typescript
await tenantStorage.updatePreferences({
  maxRecentTenants: 15
});
```

### Default Tenant

Set a default tenant to use when not on a SharePoint page:

```typescript
await tenantManager.setDefaultTenant('https://contoso.sharepoint.com');
```

## Storage Limits

The extension uses `chrome.storage.sync` with the following limits:
- **Total storage**: 100 KB
- **Max items**: 512
- **Max item size**: 8 KB

**Recommendations:**
- Keep max recent tenants ≤ 20
- Limit tenant-specific settings size
- Regularly clean up unused tenants

## Security Considerations

1. **Same-Origin Policy**: The extension validates that resolved URLs belong to the current tenant
2. **No Credential Storage**: Authentication is handled by the browser; no credentials are stored
3. **Tenant Isolation**: Settings and data are isolated per tenant
4. **URL Validation**: All URLs are validated before use

## Cloud Environment Support

### Commercial Cloud
- **Domain**: `*.sharepoint.com`
- **Example**: `https://contoso.sharepoint.com`
- **Most common deployment**

### GCC High (Government Community Cloud High)
- **Domain**: `*.sharepoint-df.com`
- **Example**: `https://contoso.sharepoint-df.com`
- **For US government agencies (DoD Impact Level 4)**

### GCC (Government Community Cloud)
- **Domain**: `*.sharepoint.us`
- **Example**: `https://contoso.sharepoint.us`
- **For US government agencies and contractors**

### China
- **Domain**: `*.sharepoint.cn`
- **Example**: `https://contoso.sharepoint.cn`
- **Operated by 21Vianet in China**

## Future Enhancements

Potential improvements for future releases:

1. **Tenant Icons**: Support custom icons/avatars for tenants
2. **Tenant Groups**: Organize tenants into groups (e.g., Clients, Internal)
3. **Quick Switcher UI**: Keyboard-driven tenant switcher
4. **Tenant-Specific Links**: Extended custom link system per tenant
5. **Import/Export**: Backup and restore tenant configurations
6. **Tenant Aliases**: Multiple display names per tenant
7. **Usage Analytics**: Track which tenants are used most frequently

## API Reference

### TenantManager

```typescript
class TenantManager {
  detectTenantFromUrl(url: string): TenantInfo | null
  getOrCreateTenant(url: string, autoRegister?: boolean): Promise<Tenant | null>
  getCurrentTenant(): Promise<Tenant | null>
  switchTenant(tenantUrl: string): Promise<TenantSwitchResult>
  getAllTenants(): Promise<Tenant[]>
  updateTenant(tenantUrl: string, updates: Partial<Tenant>): Promise<void>
  removeTenant(tenantUrl: string): Promise<void>
  setDefaultTenant(tenantUrl: string): Promise<void>
  getDefaultTenant(): Promise<Tenant | null>
  isAutoDetectEnabled(): Promise<boolean>
  setAutoDetect(enabled: boolean): Promise<void>
  validateTenantUrl(url: string): Promise<boolean>
  getTenantSettings(tenantUrl: string): Promise<Record<string, any>>
  updateTenantSettings(tenantUrl: string, settings: Record<string, any>): Promise<void>
}
```

### TenantStorage

```typescript
class TenantStorage {
  loadData(): Promise<TenantStorageData>
  saveData(data: TenantStorageData): Promise<void>
  getTenant(tenantUrl: string): Promise<Tenant | null>
  saveTenant(tenant: Tenant): Promise<void>
  removeTenant(tenantUrl: string): Promise<void>
  getAllTenants(): Promise<Tenant[]>
  getDefaultTenant(): Promise<Tenant | null>
  setDefaultTenant(tenantUrl: string): Promise<void>
  getPreferences(): Promise<TenantPreferences>
  updatePreferences(preferences: Partial<TenantPreferences>): Promise<void>
  touchTenant(tenantUrl: string): Promise<void>
  clearAll(): Promise<void>
}
```

## Related Issues

- SharePoint Context Detection (foundation for tenant detection)
- Custom Links Management (will support tenant-specific links)
- Recent/Favorite Links (will integrate with tenant context)

## Conclusion

Multi-tenant support enables SharePoint DevTools to serve users who work across multiple organizations, providing a seamless experience with automatic detection, persistent storage, and tenant-specific configurations.
