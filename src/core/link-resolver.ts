/**
 * Link Resolution Engine
 * Core engine for Issue #5: Link Resolution Engine
 *
 * Responsibilities:
 * 1. Placeholder replacement for {webUrl}, {siteUrl}, {listId}, etc.
 * 2. Link applicability logic (context-aware filtering)
 * 3. Link validation
 * 4. Final URL generation based on current context
 */

import { SharePointContext } from '../types/sharepoint-context';
import { LinkTemplate, ResolvedLink, PlaceholderType } from '../types/link-template';

/**
 * Main Link Resolver class
 */
export class LinkResolver {
  /**
   * Resolves a single link template into a resolved link
   * @param template - The link template to resolve
   * @param context - Current SharePoint context
   * @returns Resolved link or null if not applicable
   */
  static resolveLink(template: LinkTemplate, context: SharePointContext): ResolvedLink | null {
    // Check if link is applicable in current context
    if (!this.isLinkApplicable(template, context)) {
      return null;
    }

    // Replace placeholders in URL template
    const url = this.replacePlaceholders(template.urlTemplate, context);

    // Validate the resolved URL
    const validation = this.validateUrl(url, context);

    return {
      templateId: template.id,
      title: template.title,
      url,
      category: template.category,
      description: template.description,
      keywords: template.keywords,
      openInNewTab: template.openInNewTab ?? false,
      isValid: validation.isValid,
      validationErrors: validation.errors,
    };
  }

  /**
   * Resolves multiple link templates
   * @param templates - Array of link templates
   * @param context - Current SharePoint context
   * @param includeInvalid - Whether to include invalid links in results
   * @returns Array of resolved links
   */
  static resolveLinks(
    templates: LinkTemplate[],
    context: SharePointContext,
    includeInvalid = false
  ): ResolvedLink[] {
    const resolved: ResolvedLink[] = [];

    for (const template of templates) {
      const link = this.resolveLink(template, context);
      if (link && (includeInvalid || link.isValid)) {
        resolved.push(link);
      }
    }

    return resolved;
  }

  /**
   * Checks if a link template is applicable in the current context
   * @param template - The link template to check
   * @param context - Current SharePoint context
   * @returns true if link is applicable
   */
  static isLinkApplicable(template: LinkTemplate, context: SharePointContext): boolean {
    const applicability = template.applicability;

    // If not on SharePoint page, no links are applicable
    if (!context.isSharePointPage) {
      return false;
    }

    // Check page type requirements
    if (applicability.pageTypes && applicability.pageTypes.length > 0) {
      if (!applicability.pageTypes.includes(context.pageType)) {
        return false;
      }
    }

    // Check version requirements
    if (applicability.versions && applicability.versions.length > 0) {
      if (!applicability.versions.includes(context.version)) {
        return false;
      }
    }

    // Check required context fields
    if (applicability.requiredContext && applicability.requiredContext.length > 0) {
      for (const required of applicability.requiredContext) {
        if (!this.hasRequiredContext(required, context)) {
          return false;
        }
      }
    }

    return true;
  }

  /**
   * Checks if a required context field is present
   */
  private static hasRequiredContext(
    placeholder: PlaceholderType,
    context: SharePointContext
  ): boolean {
    switch (placeholder) {
      case PlaceholderType.TenantUrl:
        return context.tenantUrl !== null;
      case PlaceholderType.SiteUrl:
        return context.siteUrl !== null;
      case PlaceholderType.WebUrl:
        return context.webUrl !== null;
      case PlaceholderType.ListId:
        return context.listId !== null;
      case PlaceholderType.ListTitle:
        return context.listTitle !== null;
      default:
        return false;
    }
  }

  /**
   * Replaces placeholders in a URL template with actual values
   * @param urlTemplate - Template with placeholders (e.g., {webUrl}/_layouts/15/settings.aspx)
   * @param context - Current SharePoint context
   * @returns URL with placeholders replaced
   */
  static replacePlaceholders(urlTemplate: string, context: SharePointContext): string {
    let url = urlTemplate;

    // Define placeholder mappings
    const placeholders: Record<string, string | null> = {
      '{tenantUrl}': context.tenantUrl,
      '{siteUrl}': context.siteUrl,
      '{webUrl}': context.webUrl,
      '{listId}': context.listId,
      '{listTitle}': context.listTitle,
    };

    // Replace each placeholder
    for (const [placeholder, value] of Object.entries(placeholders)) {
      if (url.includes(placeholder)) {
        // If value is null, keep placeholder as-is (will fail validation)
        url = url.replace(new RegExp(placeholder, 'g'), value ?? placeholder);
      }
    }

    return url;
  }

  /**
   * Validates a resolved URL
   * @param url - The URL to validate
   * @param context - Current SharePoint context
   * @returns Validation result
   */
  static validateUrl(
    url: string,
    context: SharePointContext
  ): { isValid: boolean; errors: string[] } {
    const errors: string[] = [];

    // Check if placeholders remain (indicates missing context)
    if (this.hasUnresolvedPlaceholders(url)) {
      errors.push('URL contains unresolved placeholders');
    }

    // Check if URL is valid
    try {
      new URL(url);
    } catch {
      errors.push('Invalid URL format');
    }

    // Check if URL is on the same tenant (security check)
    if (context.tenantUrl && !errors.includes('Invalid URL format')) {
      try {
        const urlObj = new URL(url);
        const contextTenantObj = new URL(context.tenantUrl);

        // Allow same tenant or well-known admin URLs
        const isOwnTenant = urlObj.hostname === contextTenantObj.hostname;
        const isAdminUrl =
          urlObj.hostname.includes('admin.microsoft.com') ||
          urlObj.hostname.includes('.sharepoint.com');

        if (!isOwnTenant && !isAdminUrl) {
          errors.push('URL is not on the current tenant');
        }
      } catch {
        // Already caught in previous validation
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
    };
  }

  /**
   * Checks if a URL contains unresolved placeholders
   */
  private static hasUnresolvedPlaceholders(url: string): boolean {
    return /\{[^}]+\}/.test(url);
  }

  /**
   * Extracts all placeholders from a URL template
   * @param urlTemplate - Template string
   * @returns Array of placeholder names (without braces)
   */
  static extractPlaceholders(urlTemplate: string): string[] {
    const matches = urlTemplate.match(/\{([^}]+)\}/g);
    if (!matches) return [];

    return matches.map((match) => match.replace(/[{}]/g, ''));
  }

  /**
   * Gets missing context fields for a link template
   * @param template - Link template
   * @param context - Current SharePoint context
   * @returns Array of missing placeholder names
   */
  static getMissingContext(template: LinkTemplate, context: SharePointContext): string[] {
    const placeholders = this.extractPlaceholders(template.urlTemplate);
    const missing: string[] = [];

    for (const placeholder of placeholders) {
      // Convert placeholder name to PlaceholderType
      const placeholderType = placeholder as PlaceholderType;

      if (!this.hasRequiredContext(placeholderType, context)) {
        missing.push(placeholder);
      }
    }

    return missing;
  }

  /**
   * Filters templates to only those applicable in current context
   * @param templates - Array of link templates
   * @param context - Current SharePoint context
   * @returns Filtered array of applicable templates
   */
  static filterApplicableTemplates(
    templates: LinkTemplate[],
    context: SharePointContext
  ): LinkTemplate[] {
    return templates.filter((template) => this.isLinkApplicable(template, context));
  }

  /**
   * Groups resolved links by category
   * @param links - Array of resolved links
   * @returns Map of category to links
   */
  static groupByCategory(links: ResolvedLink[]): Map<string, ResolvedLink[]> {
    const grouped = new Map<string, ResolvedLink[]>();

    for (const link of links) {
      const category = link.category;
      if (!grouped.has(category)) {
        grouped.set(category, []);
      }
      grouped.get(category)!.push(link);
    }

    return grouped;
  }
}
