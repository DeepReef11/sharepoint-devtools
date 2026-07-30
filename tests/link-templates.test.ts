/**
 * Link Template Tests
 *
 * Guards a silent failure mode. LinkTemplateEngine resolves a template by
 * switching on the Placeholder enum; anything not in that switch returns null,
 * resolve() returns null, and link-manager drops the link from the results with
 * only a console.warn — no error, no visible symptom.
 *
 * That is how {tenantAdminUrl} removed "SharePoint Admin Center" from QuickNav
 * unnoticed. These tests fail loudly instead, without needing a live tenant.
 */
import { describe, it, expect } from '@jest/globals';
import { LinkTemplateEngine } from '../src/links/template-engine';
import { Placeholder, PlaceholderValues } from '../src/links/types';
import { SHAREPOINT_LINKS } from '../src/links/registry';

/** Every placeholder the engine can substitute. */
const SUPPORTED = new Set<string>(Object.values(Placeholder));

/** Fully-populated context — anything unresolvable here is a coverage gap. */
const FULL_CONTEXT: Required<PlaceholderValues> = {
  webUrl: 'https://contoso.sharepoint.com/sites/demo',
  siteUrl: 'https://contoso.sharepoint.com/sites/demo',
  serverUrl: 'https://contoso.sharepoint.com',
  tenantAdminUrl: 'https://contoso-admin.sharepoint.com',
  listId: '253dc79e-f232-4238-a19b-8200541ce44e',
  listUrl: 'https://contoso.sharepoint.com/sites/demo/Shared Documents',
};

describe('link template placeholders', () => {
  it('every placeholder used in the registry is supported by the engine', () => {
    const offenders: string[] = [];

    for (const link of SHAREPOINT_LINKS) {
      for (const placeholder of LinkTemplateEngine.extractPlaceholders(link.urlTemplate)) {
        if (!SUPPORTED.has(placeholder)) {
          offenders.push(`${link.id} ("${link.title}") uses ${placeholder}`);
        }
      }
    }

    expect(offenders).toEqual([]);
  });

  it('every registry link resolves against a fully-populated context', () => {
    const unresolved = SHAREPOINT_LINKS.filter(
      (link) => LinkTemplateEngine.resolve(link.urlTemplate, FULL_CONTEXT) === null
    ).map((link) => `${link.id} ("${link.title}") — ${link.urlTemplate}`);

    expect(unresolved).toEqual([]);
  });

  it('leaves no unsubstituted placeholders in a resolved URL', () => {
    for (const link of SHAREPOINT_LINKS) {
      const resolved = LinkTemplateEngine.resolve(link.urlTemplate, FULL_CONTEXT);
      expect(resolved).not.toBeNull();
      expect(resolved).not.toMatch(/\{[a-zA-Z]+\}/);
    }
  });

  it('every PlaceholderValues field has a matching enum member', () => {
    // The root cause was the type and the enum drifting apart, not the symptom.
    const fields = Object.keys(FULL_CONTEXT);
    const enumKeys = new Set(Object.values(Placeholder).map((p) => p.replace(/[{}]/g, '')));
    expect(fields.filter((f) => !enumKeys.has(f))).toEqual([]);
  });

  it('still rejects a template whose value is genuinely missing', () => {
    const resolved = LinkTemplateEngine.resolve(
      '{listUrl}/_layouts/15/listedit.aspx?List={listId}',
      { webUrl: 'https://contoso.sharepoint.com' }
    );
    expect(resolved).toBeNull();
  });

  it('resolves the tenant admin template that was previously dropped', () => {
    const resolved = LinkTemplateEngine.resolve(
      '{tenantAdminUrl}/_layouts/15/online/AdminHome.aspx',
      FULL_CONTEXT
    );
    expect(resolved).toBe(
      'https://contoso-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx'
    );
  });
});
