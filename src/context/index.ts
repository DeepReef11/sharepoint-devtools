/**
 * SharePoint Context Detection Module
 * Public API exports
 */

// Types
export {
  SharePointVersion,
  PageContext,
  SharePointUrlComponents,
  ListContext,
  SharePointContext,
  ContextDetectionResult,
} from '../types/sharepoint-context';

// URL Parser
export {
  isSharePointUrl,
  parseSharePointUrl,
  extractTenantId,
  extractListId,
  extractItemId,
  extractViewId,
  extractListName,
} from './url-parser';

// Context Detector
export {
  detectSharePointContext,
  detectCurrentContext,
  detectSharePointVersion,
  detectPageContext,
  buildListContext,
  isValidContext,
} from './context-detector';
