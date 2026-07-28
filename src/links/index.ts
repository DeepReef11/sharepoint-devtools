/**
 * SharePoint Links Registry Module
 *
 * This module provides a comprehensive registry of SharePoint navigation links
 * organized by category with support for:
 * - Template-based URLs with placeholder resolution
 * - Context-aware link filtering
 * - Priority-based link ranking
 * - Fuzzy search capabilities
 */

// Export types
export {
  Placeholder,
  LinkCategory,
  LinkContext,
  SharePointLink,
  GroupedLinks,
  PlaceholderValues,
} from './types';

// Export registry
export { SHAREPOINT_LINKS } from './registry';

// Export template engine
export { LinkTemplateEngine } from './template-engine';

// Export link manager
export { LinkManager, linkManager } from './link-manager';

// Export fuzzy search
export {
  FuzzySearchService,
  FuzzySearchOptions,
  SearchResult,
  createFuzzySearch,
} from './fuzzy-search';
