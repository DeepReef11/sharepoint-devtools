/**
 * SharePoint REST API fixtures.
 *
 * These mirror the `odata=verbose` payload shapes the extension actually consumes
 * (`data.d`, `Fields.results`, `ContentTypes.results`), so tests exercise the real
 * parsing code in ColumnFetcher rather than a stubbed version of it.
 *
 * The fixtures deliberately include values a real tenant makes tedious to produce:
 * hostile strings, every awkward field type, and empty/absent optional properties.
 */

export const TENANT = 'https://contoso.sharepoint.com';
export const WEB_URL = `${TENANT}/sites/marketing`;
export const LIST_ID = '8f2a1b3c-4d5e-6f70-8192-a3b4c5d6e7f8';

/**
 * Payloads that must never become live markup or a live link.
 * Referenced by name in assertions so a failure says what it was testing.
 */
export const HOSTILE = {
  attributeBreakout: 'x" onmouseover="alert(1)',
  singleQuoteBreakout: "x' onmouseover='alert(1)",
  scriptTag: '<script>alert(1)</script>',
  imgOnError: '<img src=x onerror="alert(1)">',
  javascriptUrl: 'javascript:alert(document.domain)',
  dataUrl: 'data:text/html,<script>alert(1)</script>',
};

interface FieldOverrides {
  [key: string]: unknown;
}

function field(
  internalName: string,
  title: string,
  typeAsString: string,
  overrides: FieldOverrides = {}
) {
  return {
    InternalName: internalName,
    StaticName: internalName,
    Title: title,
    TypeAsString: typeAsString,
    Id: `00000000-0000-0000-0000-${internalName.length.toString().padStart(12, '0')}`,
    Required: false,
    Hidden: false,
    ReadOnlyField: false,
    Description: '',
    DefaultValue: null,
    Indexed: false,
    Group: 'Custom Columns',
    ...overrides,
  };
}

/**
 * A list schema covering every branch in ColumnFetcher.parseField and
 * ColumnInspector.formatFieldValue.
 */
export const LIST_SCHEMA_RESPONSE = {
  d: {
    Title: 'Project Tasks',
    EntityTypeName: 'ProjectTasksList',
    Id: LIST_ID,
    BaseTemplate: 100,
    BaseType: 0,
    Description: 'Tracks project work',
    ContentTypesEnabled: true,
    EnableVersioning: true,
    Hidden: false,
    RootFolder: { ServerRelativeUrl: '/sites/marketing/Lists/Tasks' },
    Fields: {
      results: [
        field('Title', 'Title', 'Text', { Required: true, MaxLength: 255 }),
        field('Notes', 'Notes', 'Note'),
        field('Budget', 'Budget', 'Currency'),
        field('Estimate', 'Estimate', 'Number'),
        field('DueDate', 'Due Date', 'DateTime'),
        field('IsDone', 'Is Done', 'Boolean'),
        field('Owner', 'Owner', 'User'),
        field('Status', 'Status', 'Choice', {
          Choices: { results: ['Not Started', 'In Progress', 'Done'] },
        }),
        field('Tags', 'Tags', 'MultiChoice', {
          Choices: { results: ['alpha', 'beta'] },
        }),
        field('Department', 'Department', 'Lookup', {
          LookupList: '{11111111-2222-3333-4444-555555555555}',
          LookupField: 'Title',
        }),
        field('TotalCost', 'Total Cost', 'Calculated', {
          Formula: '=[Budget]*1.15',
        }),
        field('Homepage', 'Homepage', 'URL'),
        field('Category', 'Category', 'TaxonomyFieldType', {
          TermSetId: '99999999-8888-7777-6666-555555555555',
        }),
        field('Secret', 'Hidden Field', 'Text', { Hidden: true }),
        // Skipped by parseField's system-field filter — asserts the filter still runs.
        field('ContentTypeId', 'Content Type ID', 'Text'),
        field('owshiddenversion', 'owshiddenversion', 'Integer'),
        // Hostile metadata: a column whose *title* and *description* are attacker-chosen.
        field('Evil', HOSTILE.attributeBreakout, 'Text', {
          Description: HOSTILE.scriptTag,
        }),
      ],
    },
    ContentTypes: {
      results: [
        {
          Id: { StringValue: '0x010800ABCDEF' },
          Name: 'Task',
          Description: 'Default task content type',
        },
        {
          // Hostile content type name and description.
          Id: { StringValue: '0x010800FEDCBA' },
          Name: HOSTILE.imgOnError,
          Description: HOSTILE.attributeBreakout,
        },
      ],
    },
  },
};

/**
 * A list item whose values are attacker-controlled — this is what any site
 * contributor can store, and what the inspector renders when opened on an item.
 */
export const LIST_ITEM_RESPONSE = {
  d: {
    Id: 42,
    Title: HOSTILE.attributeBreakout,
    Notes: 'x'.repeat(250),
    Budget: 1234.5,
    Estimate: 7,
    DueDate: '2026-03-01T00:00:00Z',
    IsDone: true,
    Status: HOSTILE.scriptTag,
    Tags: { results: ['alpha', HOSTILE.imgOnError] },
    Department: { Title: HOSTILE.singleQuoteBreakout },
    TotalCost: '1419.68',
    // A URL column holding a javascript: payload — must not become a live link.
    Homepage: { Url: HOSTILE.javascriptUrl, Description: 'Click me' },
    Evil: HOSTILE.imgOnError,
  },
};

/** A URL column holding a legitimate link, for the positive case. */
export const SAFE_URL_ITEM_RESPONSE = {
  d: {
    Id: 43,
    Title: 'Ordinary item',
    Homepage: {
      Url: 'https://contoso.sharepoint.com/sites/marketing/docs',
      Description: 'Docs',
    },
  },
};
