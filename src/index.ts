/**
 * SharePoint DevTools Extension
 * Main entry point for the extension modules
 */

// Core functionality
export { LinkResolver } from './core/link-resolver';

// Context detection
export {
  detectSharePointContext,
  getCurrentContext,
  isSharePointUrl,
} from './utils/context-detector';

// Types
export type { SharePointContext, PartialSharePointContext } from './types/sharepoint-context';
export { SharePointPageType, SharePointVersion } from './types/sharepoint-context';

export type {
  LinkTemplate,
  ResolvedLink,
  LinkApplicability,
  LinkRegistry,
} from './types/link-template';
export { LinkCategory, PlaceholderType } from './types/link-template';

// Sample data
export { sampleLinkTemplates, sampleLinkRegistry } from './data/sample-links';
