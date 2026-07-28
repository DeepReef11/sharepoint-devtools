/**
 * SharePoint Permissions REST API Service
 * Provides methods to fetch SharePoint permissions data using REST APIs
 * Follows patterns established in sharepoint-api.ts
 */

/* eslint-disable @typescript-eslint/no-explicit-any */

import {
  withRetry,
  createErrorFromResponse,
  NetworkError,
  logError,
  logDebug,
} from '../utils/error-handling';
import { apiCache, generateCacheKey, CacheOptions } from './api-cache';

// =============================================================================
// Interfaces
// =============================================================================

/**
 * Represents a SharePoint user
 */
export interface SPUser {
  id: number;
  loginName: string;
  title: string;
  email: string;
  isSiteAdmin: boolean;
  principalType: number; // 1 = User, 2 = DistributionList, 4 = SecurityGroup, 8 = SharePointGroup
}

/**
 * Represents a SharePoint group
 */
export interface SPGroup {
  id: number;
  title: string;
  description: string;
  ownerTitle: string;
  allowMembersEditMembership: boolean;
  allowRequestToJoinLeave: boolean;
  onlyAllowMembersViewMembership: boolean;
  users?: SPUser[];
}

/**
 * Represents a permission level (role definition) in SharePoint
 */
export interface PermissionLevel {
  id: number;
  name: string;
  description: string;
  basePermissions: string;
  hidden: boolean;
  roleTypeKind: number; // 0 = None, 1 = Guest, 2 = Reader, 3 = Contributor, 4 = WebDesigner, 5 = Administrator
}

/**
 * Represents a role assignment (who has what permission)
 */
export interface RoleAssignment {
  principalId: number;
  principalType: 'user' | 'group';
  principalTitle: string;
  principalLoginName?: string;
  principalEmail?: string;
  roleDefinitions: PermissionLevel[];
}

/**
 * Represents the permission information for a securable object
 */
export interface PermissionInfo {
  hasUniqueRoleAssignments: boolean;
  roleAssignments: RoleAssignment[];
  scope: 'site' | 'web' | 'list' | 'item';
  scopeId?: string;
  scopeTitle?: string;
}

/**
 * Check permissions result for current user
 */
export interface UserEffectivePermissions {
  hasFullControl: boolean;
  canManagePermissions: boolean;
  canAddItems: boolean;
  canEditItems: boolean;
  canDeleteItems: boolean;
  canViewItems: boolean;
  canApproveItems: boolean;
  canManageLists: boolean;
  rawPermissions: string;
}

// =============================================================================
// Helper Functions
// =============================================================================

/**
 * Fetches data from SharePoint REST API with retry logic, error handling, and caching
 */
async function fetchPermissionsAPI<T>(url: string, cacheOptions?: CacheOptions): Promise<T> {
  if (cacheOptions) {
    const cacheKey = generateCacheKey(url);
    const cached = apiCache.get<T>(cacheKey, cacheOptions.useSessionStorage);
    if (cached !== null) {
      logDebug('Permissions API Cache Hit', { url });
      return cached;
    }
  }

  logDebug('Permissions API request', { url });

  const data = await withRetry(
    async () => {
      let response: Response;

      try {
        response = await fetch(url, {
          method: 'GET',
          headers: {
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
          },
          credentials: 'include',
        });
      } catch (error) {
        throw new NetworkError(
          'Failed to connect to SharePoint. Please check your connection.',
          error instanceof Error ? error : undefined
        );
      }

      if (!response.ok) {
        throw createErrorFromResponse(response);
      }

      try {
        const json = await response.json();
        return json.d;
      } catch (error) {
        logError('Failed to parse SharePoint permissions API response', error);
        throw new Error('Invalid response from SharePoint API');
      }
    },
    {
      maxAttempts: 3,
      initialDelayMs: 1000,
      onRetry: (attempt, error) => {
        logDebug(`Retrying permissions API request (attempt ${attempt})`, {
          url,
          error: error.message,
        });
      },
    }
  );

  if (cacheOptions) {
    const cacheKey = generateCacheKey(url);
    apiCache.set(cacheKey, data, cacheOptions);
    logDebug('Permissions API Cache Stored', { url });
  }

  return data;
}

/**
 * Maps principal type number to string
 */
function getPrincipalType(principalType: number): 'user' | 'group' {
  // 1 = User, 2 = DL, 4 = SecurityGroup, 8 = SharePointGroup
  return principalType === 1 ? 'user' : 'group';
}

/**
 * Maps role type kind to a readable name
 */
export function getRoleTypeName(roleTypeKind: number): string {
  const roleTypes: { [key: number]: string } = {
    0: 'None',
    1: 'Guest',
    2: 'Reader',
    3: 'Contributor',
    4: 'Web Designer',
    5: 'Administrator',
  };
  return roleTypes[roleTypeKind] || 'Custom';
}

// =============================================================================
// API Functions
// =============================================================================

/**
 * Gets the current user information
 * Cached for 30 minutes since user info rarely changes during a session
 */
export async function getCurrentUser(webUrl: string): Promise<SPUser> {
  const url = `${webUrl}/_api/web/currentuser`;
  const data: any = await fetchPermissionsAPI(url, {
    ttl: 30 * 60 * 1000,
    useSessionStorage: true,
  });

  return {
    id: data.Id,
    loginName: data.LoginName,
    title: data.Title,
    email: data.Email || '',
    isSiteAdmin: data.IsSiteAdmin,
    principalType: data.PrincipalType,
  };
}

/**
 * Gets all SharePoint groups in the site
 * Cached for 5 minutes since groups can change
 */
export async function getSiteGroups(webUrl: string): Promise<SPGroup[]> {
  const url = `${webUrl}/_api/web/sitegroups`;
  const data: any = await fetchPermissionsAPI(url, {
    ttl: 5 * 60 * 1000,
    useSessionStorage: true,
  });

  return data.results.map((group: any) => ({
    id: group.Id,
    title: group.Title,
    description: group.Description || '',
    ownerTitle: group.OwnerTitle || '',
    allowMembersEditMembership: group.AllowMembersEditMembership,
    allowRequestToJoinLeave: group.AllowRequestToJoinLeave,
    onlyAllowMembersViewMembership: group.OnlyAllowMembersViewMembership,
  }));
}

/**
 * Gets users in a specific SharePoint group
 * Cached for 5 minutes
 */
export async function getGroupUsers(webUrl: string, groupId: number): Promise<SPUser[]> {
  const url = `${webUrl}/_api/web/sitegroups(${groupId})/users`;
  const data: any = await fetchPermissionsAPI(url, {
    ttl: 5 * 60 * 1000,
    useSessionStorage: true,
  });

  return data.results.map((user: any) => ({
    id: user.Id,
    loginName: user.LoginName,
    title: user.Title,
    email: user.Email || '',
    isSiteAdmin: user.IsSiteAdmin,
    principalType: user.PrincipalType,
  }));
}

/**
 * Gets all permission levels (role definitions) in the site
 * Cached for 10 minutes since permission levels rarely change
 */
export async function getPermissionLevels(webUrl: string): Promise<PermissionLevel[]> {
  const url = `${webUrl}/_api/web/roledefinitions`;
  const data: any = await fetchPermissionsAPI(url, {
    ttl: 10 * 60 * 1000,
    useSessionStorage: true,
  });

  return data.results.map((role: any) => ({
    id: role.Id,
    name: role.Name,
    description: role.Description || '',
    basePermissions: role.BasePermissions?.High + ',' + role.BasePermissions?.Low,
    hidden: role.Hidden,
    roleTypeKind: role.RoleTypeKind,
  }));
}

/**
 * Gets role assignments (permissions) for the web (site)
 * Cached for 5 minutes
 */
export async function getWebPermissions(webUrl: string): Promise<PermissionInfo> {
  const baseUrl = `${webUrl}/_api/web`;

  // First check if the web has unique permissions
  const webData: any = await fetchPermissionsAPI(
    `${baseUrl}?$select=HasUniqueRoleAssignments,Title,Id`,
    {
      ttl: 5 * 60 * 1000,
      useSessionStorage: true,
    }
  );

  // Get role assignments with expanded role definitions
  const roleAssignmentsUrl = `${baseUrl}/roleassignments?$expand=Member,RoleDefinitionBindings`;
  const roleData: any = await fetchPermissionsAPI(roleAssignmentsUrl, {
    ttl: 5 * 60 * 1000,
    useSessionStorage: true,
  });

  const roleAssignments: RoleAssignment[] = roleData.results.map((ra: any) => ({
    principalId: ra.PrincipalId,
    principalType: getPrincipalType(ra.Member.PrincipalType),
    principalTitle: ra.Member.Title,
    principalLoginName: ra.Member.LoginName,
    principalEmail: ra.Member.Email,
    roleDefinitions: ra.RoleDefinitionBindings.results.map((rd: any) => ({
      id: rd.Id,
      name: rd.Name,
      description: rd.Description || '',
      basePermissions: rd.BasePermissions?.High + ',' + rd.BasePermissions?.Low,
      hidden: rd.Hidden,
      roleTypeKind: rd.RoleTypeKind,
    })),
  }));

  return {
    hasUniqueRoleAssignments: webData.HasUniqueRoleAssignments,
    roleAssignments,
    scope: 'web',
    scopeId: webData.Id,
    scopeTitle: webData.Title,
  };
}

/**
 * Gets role assignments (permissions) for a list
 * Cached for 5 minutes
 */
export async function getListPermissions(webUrl: string, listId: string): Promise<PermissionInfo> {
  const baseUrl = `${webUrl}/_api/web/lists(guid'${listId}')`;

  // First check if the list has unique permissions
  const listData: any = await fetchPermissionsAPI(
    `${baseUrl}?$select=HasUniqueRoleAssignments,Title,Id`,
    {
      ttl: 5 * 60 * 1000,
      useSessionStorage: true,
    }
  );

  // Get role assignments with expanded role definitions
  const roleAssignmentsUrl = `${baseUrl}/roleassignments?$expand=Member,RoleDefinitionBindings`;
  const roleData: any = await fetchPermissionsAPI(roleAssignmentsUrl, {
    ttl: 5 * 60 * 1000,
    useSessionStorage: true,
  });

  const roleAssignments: RoleAssignment[] = roleData.results.map((ra: any) => ({
    principalId: ra.PrincipalId,
    principalType: getPrincipalType(ra.Member.PrincipalType),
    principalTitle: ra.Member.Title,
    principalLoginName: ra.Member.LoginName,
    principalEmail: ra.Member.Email,
    roleDefinitions: ra.RoleDefinitionBindings.results.map((rd: any) => ({
      id: rd.Id,
      name: rd.Name,
      description: rd.Description || '',
      basePermissions: rd.BasePermissions?.High + ',' + rd.BasePermissions?.Low,
      hidden: rd.Hidden,
      roleTypeKind: rd.RoleTypeKind,
    })),
  }));

  return {
    hasUniqueRoleAssignments: listData.HasUniqueRoleAssignments,
    roleAssignments,
    scope: 'list',
    scopeId: listData.Id,
    scopeTitle: listData.Title,
  };
}

/**
 * Gets role assignments (permissions) for a specific list item
 * Cached for 2 minutes since item permissions can change more frequently
 */
export async function getItemPermissions(
  webUrl: string,
  listId: string,
  itemId: number
): Promise<PermissionInfo> {
  const baseUrl = `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})`;

  // First check if the item has unique permissions
  const itemData: any = await fetchPermissionsAPI(
    `${baseUrl}?$select=HasUniqueRoleAssignments,Id`,
    {
      ttl: 2 * 60 * 1000,
      useSessionStorage: true,
    }
  );

  // Get role assignments with expanded role definitions
  const roleAssignmentsUrl = `${baseUrl}/roleassignments?$expand=Member,RoleDefinitionBindings`;
  const roleData: any = await fetchPermissionsAPI(roleAssignmentsUrl, {
    ttl: 2 * 60 * 1000,
    useSessionStorage: true,
  });

  const roleAssignments: RoleAssignment[] = roleData.results.map((ra: any) => ({
    principalId: ra.PrincipalId,
    principalType: getPrincipalType(ra.Member.PrincipalType),
    principalTitle: ra.Member.Title,
    principalLoginName: ra.Member.LoginName,
    principalEmail: ra.Member.Email,
    roleDefinitions: ra.RoleDefinitionBindings.results.map((rd: any) => ({
      id: rd.Id,
      name: rd.Name,
      description: rd.Description || '',
      basePermissions: rd.BasePermissions?.High + ',' + rd.BasePermissions?.Low,
      hidden: rd.Hidden,
      roleTypeKind: rd.RoleTypeKind,
    })),
  }));

  return {
    hasUniqueRoleAssignments: itemData.HasUniqueRoleAssignments,
    roleAssignments,
    scope: 'item',
    scopeId: String(itemData.Id),
  };
}

/**
 * Gets the current user's effective permissions on the web
 * Cached for 5 minutes
 */
export async function getUserEffectivePermissionsOnWeb(
  webUrl: string
): Promise<UserEffectivePermissions> {
  const url = `${webUrl}/_api/web/effectivebasepermissions`;
  const data: any = await fetchPermissionsAPI(url, {
    ttl: 5 * 60 * 1000,
    useSessionStorage: true,
  });

  // SharePoint returns High and Low values for 64-bit permission mask
  const high = data.High;
  const low = data.Low;

  return parseEffectivePermissions(high, low);
}

/**
 * Gets the current user's effective permissions on a list
 * Cached for 5 minutes
 */
export async function getUserEffectivePermissionsOnList(
  webUrl: string,
  listId: string
): Promise<UserEffectivePermissions> {
  const url = `${webUrl}/_api/web/lists(guid'${listId}')/effectivebasepermissions`;
  const data: any = await fetchPermissionsAPI(url, {
    ttl: 5 * 60 * 1000,
    useSessionStorage: true,
  });

  const high = data.High;
  const low = data.Low;

  return parseEffectivePermissions(high, low);
}

/**
 * Parses the effective permissions bitmask
 */
function parseEffectivePermissions(high: number, low: number): UserEffectivePermissions {
  // SharePoint permission flags (simplified subset)
  // Full documentation: https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-csom/ee536458(v=office.15)
  const hasPermission = (mask: number, isHigh: boolean): boolean => {
    return isHigh ? (high & mask) === mask : (low & mask) === mask;
  };

  return {
    // FullMask (high: 2147483647)
    hasFullControl: high === 2147483647,
    // ManagePermissions (high: 32)
    canManagePermissions: hasPermission(32, true),
    // AddListItems (low: 2)
    canAddItems: hasPermission(2, false),
    // EditListItems (low: 4)
    canEditItems: hasPermission(4, false),
    // DeleteListItems (low: 8)
    canDeleteItems: hasPermission(8, false),
    // ViewListItems (low: 1)
    canViewItems: hasPermission(1, false),
    // ApproveItems (low: 16)
    canApproveItems: hasPermission(16, false),
    // ManageLists (high: 2048)
    canManageLists: hasPermission(2048, true),
    rawPermissions: `${high},${low}`,
  };
}

/**
 * Generates URL for the site permissions page
 */
export function getSitePermissionsUrl(webUrl: string): string {
  return `${webUrl}/_layouts/15/user.aspx`;
}

/**
 * Generates URL for the list permissions page
 */
export function getListPermissionsUrl(webUrl: string, listId: string): string {
  return `${webUrl}/_layouts/15/user.aspx?List={${listId}}`;
}

/**
 * Generates URL for managing a specific group
 */
export function getGroupManagementUrl(webUrl: string, groupId: number): string {
  return `${webUrl}/_layouts/15/people.aspx?MembershipGroupId=${groupId}`;
}
