/**
 * SharePoint Online cloud environments.
 *
 * Single source of truth for "is this a SharePoint host, and which cloud is it?".
 * This logic was previously copy-pasted across eight call sites, which is how the
 * manifest and the detection code drifted out of sync with each other.
 *
 * Domains per Microsoft's published endpoint lists:
 *   Commercial      *.sharepoint.com
 *   GCC             *.sharepoint.com      (GCC runs on the worldwide endpoints)
 *   GCC High        *.sharepoint.us
 *   DoD             *.sharepoint-mil.us
 *   China/21Vianet  *.sharepoint.cn
 *   Dogfood ring    *.sharepoint-df.com   (Microsoft-internal first-release)
 *
 * Note that GCC is NOT distinguishable from Commercial by hostname alone — both
 * live on *.sharepoint.com. Anything needing to tell them apart has to ask the
 * tenant, not the URL.
 */

import { SharePointCloud } from '../types/tenant';

/** Suffixes that identify a SharePoint Online host, longest-first. */
const CLOUD_SUFFIXES: { suffix: string; cloud: SharePointCloud }[] = [
  // -mil.us and -df.com must be tested before the bare .us / .com suffixes.
  { suffix: '.sharepoint-mil.us', cloud: SharePointCloud.DoD },
  { suffix: '.sharepoint-df.com', cloud: SharePointCloud.Dogfood },
  { suffix: '.sharepoint.us', cloud: SharePointCloud.GCCHigh },
  { suffix: '.sharepoint.cn', cloud: SharePointCloud.China },
  { suffix: '.sharepoint.com', cloud: SharePointCloud.Commercial },
];

/**
 * Host match patterns for the extension manifest.
 * Keep `public/manifest.json` in sync with this list — a cloud that is detected
 * here but missing from the manifest can never be reached, because the content
 * script is not injected on it.
 */
export const SHAREPOINT_HOST_PATTERNS = [
  'https://*.sharepoint.com/*',
  'https://*.sharepoint-df.com/*',
  'https://*.sharepoint.us/*',
  'https://*.sharepoint-mil.us/*',
  'https://*.sharepoint.cn/*',
];

/** Returns true when the hostname belongs to any SharePoint Online cloud. */
export function isSharePointHostname(hostname: string): boolean {
  const host = hostname.toLowerCase();
  return CLOUD_SUFFIXES.some(({ suffix }) => host.endsWith(suffix));
}

/** Returns true when the URL points at any SharePoint Online cloud. */
export function isSharePointUrl(url: string): boolean {
  try {
    return isSharePointHostname(new URL(url).hostname);
  } catch {
    return false;
  }
}

/**
 * Identifies the cloud a hostname belongs to.
 * Commercial is returned for GCC tenants too — see the note above.
 */
export function detectCloud(hostname: string): SharePointCloud {
  const host = hostname.toLowerCase();
  const match = CLOUD_SUFFIXES.find(({ suffix }) => host.endsWith(suffix));
  return match ? match.cloud : SharePointCloud.Unknown;
}

/** Extracts the tenant prefix, e.g. "contoso" from "contoso.sharepoint.com". */
export function extractTenantName(hostname: string): string {
  const host = hostname.toLowerCase();
  const prefix = host.split('.')[0];
  // Admin hosts are "<tenant>-admin"; strip the suffix so both forms agree.
  return prefix.endsWith('-admin') ? prefix.slice(0, -'-admin'.length) : prefix;
}

/**
 * Builds the tenant admin centre URL for any cloud, e.g.
 * "contoso.sharepoint.us" -> "https://contoso-admin.sharepoint.us".
 * Returns an empty string for non-SharePoint hosts.
 */
export function getTenantAdminUrl(hostname: string): string {
  const host = hostname.toLowerCase();
  const match = CLOUD_SUFFIXES.find(({ suffix }) => host.endsWith(suffix));
  if (!match) return '';

  return `https://${extractTenantName(host)}-admin${match.suffix}`;
}
