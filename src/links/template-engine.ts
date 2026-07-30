import { Placeholder, PlaceholderValues } from './types';

/**
 * Template engine for resolving SharePoint link placeholders
 */
export class LinkTemplateEngine {
  /**
   * Replaces placeholders in a URL template with actual values
   * @param template - URL template with placeholders
   * @param values - Placeholder values from current context
   * @returns Resolved URL or null if required placeholders are missing
   */
  static resolve(template: string, values: PlaceholderValues): string | null {
    let resolvedUrl = template;

    // Check if template contains placeholders
    const placeholders = this.extractPlaceholders(template);

    // Validate that all required placeholder values are available
    for (const placeholder of placeholders) {
      const value = this.getPlaceholderValue(placeholder, values);

      if (value === null || value === undefined) {
        // Missing required placeholder value
        return null;
      }

      resolvedUrl = resolvedUrl.replace(placeholder, value);
    }

    return resolvedUrl;
  }

  /**
   * Extracts all placeholders from a template string
   * @param template - URL template
   * @returns Array of placeholder strings
   */
  static extractPlaceholders(template: string): string[] {
    const placeholderRegex = /\{[^}]+\}/g;
    return template.match(placeholderRegex) || [];
  }

  /**
   * Gets the value for a specific placeholder
   * @param placeholder - Placeholder string (e.g., '{webUrl}')
   * @param values - Available placeholder values
   * @returns The value or null if not found
   */
  private static getPlaceholderValue(
    placeholder: string,
    values: PlaceholderValues
  ): string | null {
    switch (placeholder) {
      case Placeholder.WebUrl:
        return values.webUrl || null;
      case Placeholder.SiteUrl:
        return values.siteUrl || null;
      case Placeholder.ServerUrl:
        return values.serverUrl || null;
      case Placeholder.TenantAdminUrl:
        return values.tenantAdminUrl || null;
      case Placeholder.ListId:
        return values.listId || null;
      case Placeholder.ListUrl:
        return values.listUrl || null;
      default:
        console.warn(`Unknown placeholder: ${placeholder}`);
        return null;
    }
  }

  /**
   * Checks if a template can be resolved with given values
   * @param template - URL template
   * @param values - Available placeholder values
   * @returns True if template can be resolved
   */
  static canResolve(template: string, values: PlaceholderValues): boolean {
    const placeholders = this.extractPlaceholders(template);

    return placeholders.every((placeholder) => {
      const value = this.getPlaceholderValue(placeholder, values);
      return value !== null && value !== undefined;
    });
  }

  /**
   * Returns list of missing placeholders for a template
   * @param template - URL template
   * @param values - Available placeholder values
   * @returns Array of missing placeholder names
   */
  static getMissingPlaceholders(template: string, values: PlaceholderValues): string[] {
    const placeholders = this.extractPlaceholders(template);

    return placeholders.filter((placeholder) => {
      const value = this.getPlaceholderValue(placeholder, values);
      return value === null || value === undefined;
    });
  }
}
