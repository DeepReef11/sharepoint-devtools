/**
 * @jest-environment-options {"url": "https://contoso.sharepoint.com/sites/marketing/Lists/Tasks/AllItems.aspx"}
 */

/**
 * Permission Inspector integration tests.
 *
 * Covers a keyboard trap that shipped: the "Manage Permissions" link sits in the
 * info banner, outside any `.sp-perm-card`, so the card navigation model never
 * focuses it and Tab is the only way to reach it. The document-level keydown
 * handler nonetheless claimed Enter and called preventDefault, so activating the
 * focused link did nothing and the panel silently entered inner-navigation on
 * whichever card happened to be selected.
 */

import { describe, it, expect, beforeEach, afterEach } from '@jest/globals';
import { PermissionInspector } from '../src/ui/permission-inspector';
import { installSharePointPage, type ApiRoute } from './helpers/sharepoint-env';

const WEB_ROUTE: ApiRoute = {
  match: '$select=HasUniqueRoleAssignments',
  body: {
    d: {
      HasUniqueRoleAssignments: true,
      Title: 'Marketing',
      Id: '31a6898e-565e-42c3-a960-c7b6e5402973',
    },
  },
};

const ROLE_ASSIGNMENTS_ROUTE: ApiRoute = {
  match: '/roleassignments',
  body: {
    d: {
      results: [
        {
          PrincipalId: 3,
          Member: {
            PrincipalType: 8,
            Title: 'Marketing Owners',
            LoginName: 'c:0o.c|federateddirectoryclaimprovider|owners',
            Email: '',
          },
          RoleDefinitionBindings: {
            results: [
              { Id: 1073741829, Name: 'Full Control', Description: 'Has full control.', Hidden: false, RoleTypeKind: 5 },
            ],
          },
        },
        {
          PrincipalId: 5,
          Member: {
            PrincipalType: 8,
            Title: 'Marketing Members',
            LoginName: 'c:0o.c|federateddirectoryclaimprovider|members',
            Email: '',
          },
          RoleDefinitionBindings: {
            results: [
              { Id: 1073741830, Name: 'Edit', Description: 'Can add, edit and delete.', Hidden: false, RoleTypeKind: 6 },
            ],
          },
        },
      ],
    },
  },
};

const GROUPS_ROUTE: ApiRoute = { match: '/sitegroups', body: { d: { results: [] } } };
const LEVELS_ROUTE: ApiRoute = { match: '/roledefinitions', body: { d: { results: [] } } };
const USER_ROUTE: ApiRoute = {
  match: '/currentuser',
  body: { d: { Id: 15, Title: 'Klein Bottle', LoginName: 'i:0#.f|membership|klein', IsSiteAdmin: true } },
};

// First substring match wins, so the narrow routes precede the broad ones.
const ROUTES = [WEB_ROUTE, ROLE_ASSIGNMENTS_ROUTE, GROUPS_ROUTE, LEVELS_ROUTE, USER_ROUTE];

let teardown: () => void;

async function openInspector() {
  teardown = installSharePointPage({ routes: ROUTES });
  const inspector = new PermissionInspector();
  await inspector.show();
  return inspector;
}

function pressEnterOn(target: HTMLElement): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
  target.dispatchEvent(event);
  return event;
}

beforeEach(() => {
  // Cached permission responses would leak between tests.
  sessionStorage.clear();
  // jsdom implements no scrolling; the panel scrolls the selected card into
  // view after every handled key, and the resulting TypeError is thrown inside
  // the listener where it is easy to mistake for a product failure.
  Element.prototype.scrollIntoView = function scrollIntoView() {
    /* no-op for jsdom */
  };
});

afterEach(() => {
  teardown?.();
  document.getElementById('sp-permission-inspector')?.remove();
  document.body.innerHTML = '';
});

describe('PermissionInspector keyboard handling', () => {
  it('renders role assignment cards and a Manage Permissions link', async () => {
    await openInspector();

    const cards = document.querySelectorAll('.sp-perm-card');
    expect(cards.length).toBeGreaterThan(0);

    const link = document.querySelector('a.sp-perm-link-btn') as HTMLAnchorElement;
    expect(link).not.toBeNull();
    expect(link.textContent).toMatch(/Manage Permissions/i);
  });

  it('places the Manage Permissions link outside any card', async () => {
    await openInspector();

    // This is why the card navigation model cannot reach it, and why Tab plus a
    // working Enter is the only keyboard route to it.
    const link = document.querySelector('a.sp-perm-link-btn') as HTMLAnchorElement;
    expect(link.closest('.sp-perm-card')).toBeNull();
  });

  it('lets Enter activate the focused Manage Permissions link', async () => {
    await openInspector();

    const link = document.querySelector('a.sp-perm-link-btn') as HTMLAnchorElement;
    link.focus();

    const event = pressEnterOn(link);

    // The handler must stand aside so the browser performs the navigation.
    expect(event.defaultPrevented).toBe(false);
    // And it must not have hijacked the key to enter inner-navigation instead.
    expect(document.querySelector('.sp-perm-card.inner-nav-active')).toBeNull();
  });

  it('still claims Enter when focus is not on a control', async () => {
    await openInspector();

    // The panel keeps ownership of Enter everywhere else, so the fix is a
    // carve-out for real controls rather than a blanket surrender of the key.
    // (Whether inner-navigation then activates depends on the card having
    // focusable children — assignment cards have none, group cards do.)
    const panel = document.getElementById('sp-permission-inspector') as HTMLElement;
    const event = pressEnterOn(panel);

    expect(event.defaultPrevented).toBe(true);
  });

  it('closes on Escape when the filter matches nothing', async () => {
    const inspector = await openInspector();

    // Filter to something no principal matches, so the card list empties.
    const search = document.getElementById('sp-perm-inspector-search') as HTMLInputElement;
    search.value = 'zzzz-no-such-principal';
    search.dispatchEvent(new Event('input', { bubbles: true }));
    await new Promise((r) => setTimeout(r, 50));
    expect(document.querySelectorAll('#sp-perm-inspector-data .sp-perm-card').length).toBe(0);

    // With no cards the handler used to return before reaching the Escape case,
    // leaving the panel open with no keyboard way out.
    const panel = document.getElementById('sp-permission-inspector') as HTMLElement;
    panel.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));

    expect(panel.style.display).toBe('none');
    expect(inspector).toBeDefined();
  });

  it('does not hijack Space on the focused link either', async () => {
    await openInspector();

    const link = document.querySelector('a.sp-perm-link-btn') as HTMLAnchorElement;
    link.focus();

    const event = new KeyboardEvent('keydown', { key: ' ', bubbles: true, cancelable: true });
    link.dispatchEvent(event);

    expect(event.defaultPrevented).toBe(false);
  });
});
