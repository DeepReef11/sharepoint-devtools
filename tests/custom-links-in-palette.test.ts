/**
 * Custom links reaching the palette.
 *
 * CustomLinksStorage was read and written only by the options page. A link the
 * user created was stored, listed and editable there, and never appeared in
 * QuickNav — the one place it existed for. LinkManager is what the palette
 * searches, so the merge belongs there and is covered here.
 */

import { describe, it, expect, beforeEach } from '@jest/globals';
import { LinkManager } from '../src/links/link-manager';
import { createFuzzySearch } from '../src/links/fuzzy-search';
import { SHAREPOINT_LINKS } from '../src/links/registry';
import { SharePointLink, LinkCategory } from '../src/links/types';

const custom = (over: Partial<SharePointLink> = {}): SharePointLink => ({
  id: 'my-team-dashboard',
  title: 'Team Dashboard',
  urlTemplate: '{webUrl}/SitePages/Dashboard.aspx',
  category: LinkCategory.Content,
  description: 'The dashboard the team actually uses',
  ...over,
});

// Everything the built-in registry needs resolves from these.
const VALUES = {
  webUrl: 'https://contoso.sharepoint.com/sites/marketing',
  siteUrl: 'https://contoso.sharepoint.com/sites/marketing',
  serverUrl: 'https://contoso.sharepoint.com',
  tenantAdminUrl: 'https://contoso-admin.sharepoint.com',
};

let manager: LinkManager;

beforeEach(() => {
  manager = new LinkManager();
});

describe('LinkManager.setCustomLinks', () => {
  it('adds custom links to the searchable set', () => {
    expect(manager.getLinkById('my-team-dashboard')).toBeUndefined();

    manager.setCustomLinks([custom()]);

    expect(manager.getLinkById('my-team-dashboard')?.title).toBe('Team Dashboard');
  });

  it('keeps the built-in registry alongside them', () => {
    manager.setCustomLinks([custom()]);

    expect(manager.getAllLinks().length).toBe(SHAREPOINT_LINKS.length + 1);
    expect(manager.getLinkById('site-contents')).toBeDefined();
  });

  it('returns them from getApplicableLinks, which is what the palette renders', () => {
    manager.setCustomLinks([custom()]);

    const applicable = manager.getApplicableLinks(VALUES, false);
    expect(applicable.map((l) => l.id)).toContain('my-team-dashboard');
  });

  it('makes them findable by fuzzy search', () => {
    manager.setCustomLinks([custom()]);

    const search = createFuzzySearch(manager.getApplicableLinks(VALUES, false));
    const titles = search.search('dashboard', { limit: 20 }).map((r) => r.item.title);

    expect(titles).toContain('Team Dashboard');
  });

  it('replaces the previous set rather than accumulating', () => {
    manager.setCustomLinks([custom()]);
    manager.setCustomLinks([custom({ id: 'other', title: 'Other' })]);

    // The user deleted the first one on the options page; it must go.
    expect(manager.getLinkById('my-team-dashboard')).toBeUndefined();
    expect(manager.getLinkById('other')).toBeDefined();
    expect(manager.getAllLinks().length).toBe(SHAREPOINT_LINKS.length + 1);
  });

  it('reflects an edit to an existing custom link', () => {
    manager.setCustomLinks([custom()]);
    manager.setCustomLinks([custom({ title: 'Renamed Dashboard' })]);

    expect(manager.getLinkById('my-team-dashboard')?.title).toBe('Renamed Dashboard');
    expect(manager.getAllLinks().length).toBe(SHAREPOINT_LINKS.length + 1);
  });

  it('lets a custom link override a built-in with the same id', () => {
    const builtIn = SHAREPOINT_LINKS[0];

    manager.setCustomLinks([custom({ id: builtIn.id, title: 'My Version' })]);

    const matches = manager.getAllLinks().filter((l) => l.id === builtIn.id);
    expect(matches).toHaveLength(1);
    expect(matches[0].title).toBe('My Version');
  });

  it('clears the set when the user deletes their last link', () => {
    manager.setCustomLinks([custom()]);
    manager.setCustomLinks([]);

    expect(manager.getLinkById('my-team-dashboard')).toBeUndefined();
    expect(manager.getAllLinks().length).toBe(SHAREPOINT_LINKS.length);
  });
});

describe('LinkManager.isCustomLink', () => {
  it('distinguishes user links from built-ins', () => {
    manager.setCustomLinks([custom()]);

    expect(manager.isCustomLink('my-team-dashboard')).toBe(true);
    expect(manager.isCustomLink('site-contents')).toBe(false);
  });

  it('is false before any custom links are loaded', () => {
    expect(manager.isCustomLink('my-team-dashboard')).toBe(false);
  });

  // The default (empty query) list keeps only priority >= 5, to cut 75-odd
  // built-ins down to something readable. The options page makes priority
  // optional, so without this flag a custom link saved without one would be
  // searchable but absent from the list the palette opens on.
  it('identifies a custom link that carries no priority', () => {
    manager.setCustomLinks([custom()]);

    const link = manager.getLinkById('my-team-dashboard')!;
    expect(link.priority).toBeUndefined();
    expect(manager.isCustomLink(link.id) || (link.priority || 0) >= 5).toBe(true);
  });
});
