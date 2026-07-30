/**
 * Content definitions for the dev-tenant seeder.
 *
 * Field types are chosen deliberately to cover everything the column and object
 * inspectors have to render: Choice, MultiChoice, User, Lookup, Currency,
 * Number, DateTime, Boolean, Note, URL and Calculated. If a type is missing
 * here, it is untested there.
 */

// Dates are fixed rather than relative so re-runs stay comparable.
const d = (s) => `${s}T12:00:00Z`;

/**
 * Principals that person columns and group membership are assigned to.
 *
 * These are deliberately fictional. Filling person columns from whichever real
 * accounts the tenant happened to contain put colleagues' names into demo data
 * and, from there, into screen recordings committed to this repo. The seeder
 * creates nothing in the directory, so these accounts have to exist first — see
 * tools/README.md. Any that cannot be resolved are skipped, and if none resolve
 * the seeder falls back to the account it is running as.
 */
const PERSONAS = [
  { alias: 'dana.reyes', name: 'Dana Reyes' },
  { alias: 'marcus.webb', name: 'Marcus Webb' },
  { alias: 'priya.nair', name: 'Priya Nair' },
];

const GROUPS = [
  {
    title: 'Project Managers',
    description: 'Owns delivery schedules and budgets',
    role: 'Edit',
    members: [0],
  },
  {
    title: 'Developers',
    description: 'Build team — contribute to backlog and docs',
    role: 'Contribute',
    members: [0, 1, 2],
    claimMembers: ['Everyone except external users'],
  },
  {
    title: 'QA Reviewers',
    description: 'Read-only access for verification passes',
    role: 'Read',
    members: [1, 2],
    // A claims principal makes group membership look like a real tenant's and
    // gives the permission inspector's Groups tab something to expand into.
    claimMembers: ['Everyone except external users'],
  },
  {
    title: 'Contract Auditors',
    description: 'Restricted group used to demonstrate unique permissions',
    role: 'Read',
    members: [2],
  },
];

const CHOICE = (name, display, choices, def) =>
  `<Field Type="Choice" DisplayName="${display}" Name="${name}" Format="Dropdown">` +
  `<CHOICES>${choices.map((c) => `<CHOICE>${c}</CHOICE>`).join('')}</CHOICES>` +
  (def ? `<Default>${def}</Default>` : '') +
  `</Field>`;

const LISTS = [
  // -------------------------------------------------------------------------
  {
    title: 'Projects',
    description: 'Active engagements, budgets and owners',
    template: 100,
    fields: [
      { name: 'ProjectCode', xml: '<Field Type="Text" DisplayName="Project Code" Name="ProjectCode" MaxLength="20" />' },
      { name: 'Status', xml: CHOICE('Status', 'Status', ['Not Started', 'In Progress', 'Blocked', 'Complete'], 'Not Started') },
      { name: 'Phase', xml: CHOICE('Phase', 'Phase', ['Discovery', 'Build', 'UAT', 'Launch']) },
      { name: 'Owner', xml: '<Field Type="User" DisplayName="Owner" Name="Owner" UserSelectionMode="PeopleOnly" />' },
      { name: 'StartDate', xml: '<Field Type="DateTime" DisplayName="Start Date" Name="StartDate" Format="DateOnly" />' },
      { name: 'DueDate', xml: '<Field Type="DateTime" DisplayName="Due Date" Name="DueDate" Format="DateOnly" />' },
      { name: 'Budget', xml: '<Field Type="Currency" DisplayName="Budget" Name="Budget" Decimals="2" LCID="1033" />' },
      { name: 'PercentComplete', xml: '<Field Type="Number" DisplayName="Percent Complete" Name="PercentComplete" Min="0" Max="100" />' },
      { name: 'IsBillable', xml: '<Field Type="Boolean" DisplayName="Billable" Name="IsBillable"><Default>1</Default></Field>' },
      { name: 'Summary', xml: '<Field Type="Note" DisplayName="Summary" Name="Summary" NumLines="6" RichText="FALSE" />' },
      { name: 'SpecLink', xml: '<Field Type="URL" DisplayName="Spec Link" Name="SpecLink" Format="Hyperlink" />' },
    ],
    rows: ({ humans }) => {
      const owner = (i) => (humans[i % humans.length] ? { OwnerId: humans[i % humans.length].Id } : {});
      const rows = [
        ['PRJ-1001', 'Intranet Redesign', 'In Progress', 'Build', '2026-02-02', '2026-08-28', 148000, 62, true,
          'Replace the legacy publishing portal with a modern communication site. Includes IA rework and a component library.'],
        ['PRJ-1002', 'Document Migration', 'In Progress', 'Build', '2026-03-16', '2026-09-30', 96500, 41, true,
          'Move 2.4TB of file-share content into SharePoint with metadata mapping and retention labels applied.'],
        ['PRJ-1003', 'Records Retention Policy', 'Blocked', 'Discovery', '2026-01-12', '2026-07-31', 42000, 18, false,
          'Blocked pending legal sign-off on the retention schedule for finance records.'],
        ['PRJ-1004', 'Teams Governance Rollout', 'Not Started', 'Discovery', '2026-09-01', '2027-01-15', 61000, 0, true,
          'Naming conventions, expiry policy and provisioning workflow for Teams.'],
        ['PRJ-1005', 'Search Relevance Tuning', 'Complete', 'Launch', '2025-10-06', '2026-03-27', 33500, 100, false,
          'Managed property mapping and query rules; measured a 22% improvement in click-through.'],
        ['PRJ-1006', 'Power Platform CoE', 'In Progress', 'UAT', '2026-04-20', '2026-11-13', 87250, 55, true,
          'Centre of excellence toolkit, environment strategy and DLP policy baseline.'],
        ['PRJ-1007', 'Extranet for Partners', 'Blocked', 'Build', '2026-05-11', '2026-12-18', 112000, 27, true,
          'External sharing model with B2B guest access. Blocked on conditional access review.'],
        ['PRJ-1008', 'Accessibility Audit', 'Not Started', 'Discovery', '2026-10-05', '2027-02-26', 28000, 0, false,
          'WCAG 2.2 AA audit across the top 50 pages, with a remediation backlog.'],
      ];
      return rows.map((r, i) => ({
        Title: r[1], ProjectCode: r[0], Status: r[2], Phase: r[3],
        StartDate: d(r[4]), DueDate: d(r[5]), Budget: r[6],
        PercentComplete: r[7], IsBillable: r[8], Summary: r[9],
        SpecLink: { __metadata: { type: 'SP.FieldUrlValue' }, Url: `https://contoso.sharepoint.com/specs/${r[0]}`, Description: `${r[0]} spec` },
        ...owner(i),
      }));
    },
  },

  // -------------------------------------------------------------------------
  {
    title: 'Tasks Backlog',
    description: 'Work items linked to projects',
    template: 100,
    fields: [
      { name: 'AssignedTo', xml: '<Field Type="User" DisplayName="Assigned To" Name="AssignedTo" UserSelectionMode="PeopleOnly" />' },
      { name: 'Status', xml: CHOICE('Status', 'Status', ['Backlog', 'In Progress', 'In Review', 'Done'], 'Backlog') },
      { name: 'Priority', xml: CHOICE('Priority', 'Priority', ['Low', 'Normal', 'High', 'Critical'], 'Normal') },
      { name: 'Estimate', xml: '<Field Type="Number" DisplayName="Estimate (h)" Name="Estimate" Min="0" />' },
      { name: 'DueDate', xml: '<Field Type="DateTime" DisplayName="Due Date" Name="DueDate" Format="DateOnly" />' },
      {
        name: 'ProjectRef',
        xml: (ids) =>
          ids.Projects
            ? `<Field Type="Lookup" DisplayName="Project" Name="ProjectRef" List="{${ids.Projects}}" ShowField="Title" />`
            : null,
      },
    ],
    rows: async ({ sp, humans }) => {
      const projects = await sp("/_api/web/lists/getbytitle('Projects')/items?$select=Id,Title&$top=50");
      const byTitle = Object.fromEntries((projects?.results || []).map((p) => [p.Title, p.Id]));
      const who = (i) => (humans[i % humans.length] ? { AssignedToId: humans[i % humans.length].Id } : {});
      const rows = [
        ['Audit existing page templates', 'Done', 'Normal', 8, '2026-03-06', 'Intranet Redesign'],
        ['Build hub navigation component', 'In Progress', 'High', 21, '2026-08-07', 'Intranet Redesign'],
        ['Define metadata crosswalk', 'In Review', 'High', 13, '2026-06-19', 'Document Migration'],
        ['Dry-run migration for Finance', 'In Progress', 'Critical', 34, '2026-07-24', 'Document Migration'],
        ['Draft retention schedule', 'Backlog', 'Normal', 13, '2026-07-10', 'Records Retention Policy'],
        ['Publish naming convention guide', 'Backlog', 'Low', 5, '2026-11-06', 'Teams Governance Rollout'],
        ['Map managed properties', 'Done', 'Normal', 8, '2026-01-30', 'Search Relevance Tuning'],
        ['Baseline DLP policies', 'In Progress', 'High', 21, '2026-09-11', 'Power Platform CoE'],
        ['Guest access threat model', 'Backlog', 'Critical', 13, '2026-10-02', 'Extranet for Partners'],
        ['Contrast audit on top 50 pages', 'Backlog', 'Normal', 8, '2027-01-08', 'Accessibility Audit'],
        ['Keyboard traps in nav flyout', 'Backlog', 'High', 5, '2027-01-22', 'Accessibility Audit'],
        ['Retire legacy web parts', 'In Review', 'Normal', 13, '2026-08-21', 'Intranet Redesign'],
      ];
      return rows.map((r, i) => ({
        Title: r[0], Status: r[1], Priority: r[2], Estimate: r[3], DueDate: d(r[4]),
        ...(byTitle[r[5]] ? { ProjectRefId: byTitle[r[5]] } : {}),
        ...who(i),
      }));
    },
  },

  // -------------------------------------------------------------------------
  {
    title: 'Issues',
    description: 'Defects raised against delivered work',
    template: 100,
    fields: [
      { name: 'Severity', xml: CHOICE('Severity', 'Severity', ['Trivial', 'Minor', 'Major', 'Blocker'], 'Minor') },
      { name: 'Component', xml: CHOICE('Component', 'Component', ['Navigation', 'Search', 'Permissions', 'Migration', 'Rendering']) },
      { name: 'Environments', xml: '<Field Type="MultiChoice" DisplayName="Environments" Name="Environments"><CHOICES><CHOICE>Dev</CHOICE><CHOICE>UAT</CHOICE><CHOICE>Prod</CHOICE></CHOICES></Field>' },
      { name: 'ReportedBy', xml: '<Field Type="User" DisplayName="Reported By" Name="ReportedBy" UserSelectionMode="PeopleOnly" />' },
      { name: 'Repro', xml: '<Field Type="Note" DisplayName="Repro Steps" Name="Repro" NumLines="8" RichText="FALSE" />' },
      { name: 'IsResolved', xml: '<Field Type="Boolean" DisplayName="Resolved" Name="IsResolved"><Default>0</Default></Field>' },
      { name: 'FoundOn', xml: '<Field Type="DateTime" DisplayName="Found On" Name="FoundOn" Format="DateOnly" />' },
      { name: 'TicketLink', xml: '<Field Type="URL" DisplayName="Ticket" Name="TicketLink" Format="Hyperlink" />' },
    ],
    rows: ({ humans }) => {
      const by = (i) => (humans[i % humans.length] ? { ReportedById: humans[i % humans.length].Id } : {});
      const rows = [
        ['Breadcrumb drops the hub level', 'Major', 'Navigation', 'UAT', false, '2026-06-11',
          '1. Open a subsite page from the hub\n2. Observe breadcrumb\nExpected: hub > site > page\nActual: site > page'],
        ['Search returns stale titles after rename', 'Minor', 'Search', 'Prod', false, '2026-05-28',
          '1. Rename a page\n2. Search for the new title\nExpected: new title\nActual: previous title for ~15 minutes'],
        ['Unique permissions not copied on move', 'Blocker', 'Permissions', 'UAT;#Prod', false, '2026-07-02',
          '1. Break inheritance on a document\n2. Move it to another library\nExpected: unique permissions retained\nActual: inherits destination'],
        ['Metadata lost for .msg attachments', 'Major', 'Migration', 'Dev;#UAT', true, '2026-04-17',
          '1. Migrate a folder containing .msg files\n2. Inspect the sent date column\nExpected: populated\nActual: empty'],
        ['Column widths reset on refresh', 'Trivial', 'Rendering', 'Dev', true, '2026-03-09',
          '1. Resize a list column\n2. Refresh\nExpected: width persists\nActual: resets to default'],
        ['People picker slow above 5k users', 'Minor', 'Permissions', 'Prod', false, '2026-06-30',
          '1. Type three characters in the people picker\nExpected: results under 1s\nActual: 4-6s'],
        ['Modern list view breaks in narrow viewport', 'Major', 'Rendering', 'UAT', false, '2026-07-15',
          '1. Open a list at 360px width\nExpected: responsive layout\nActual: horizontal overflow clips the command bar'],
      ];
      return rows.map((r, i) => ({
        Title: r[0], Severity: r[1], Component: r[2],
        Environments: { __metadata: { type: 'Collection(Edm.String)' }, results: r[3].split(';#') },
        IsResolved: r[4], FoundOn: d(r[5]), Repro: r[6],
        TicketLink: { __metadata: { type: 'SP.FieldUrlValue' }, Url: `https://contoso.example/tickets/${1200 + i}`, Description: `TICKET-${1200 + i}` },
        ...by(i),
      }));
    },
  },

  // -------------------------------------------------------------------------
  {
    title: 'Equipment',
    description: 'Loaned hardware and asset tracking',
    template: 100,
    fields: [
      { name: 'AssetTag', xml: '<Field Type="Text" DisplayName="Asset Tag" Name="AssetTag" MaxLength="30" />' },
      { name: 'Category', xml: CHOICE('Category', 'Category', ['Laptop', 'Monitor', 'Phone', 'Peripheral', 'Server']) },
      { name: 'AssignedUser', xml: '<Field Type="User" DisplayName="Assigned To" Name="AssignedUser" UserSelectionMode="PeopleOnly" />' },
      { name: 'PurchaseDate', xml: '<Field Type="DateTime" DisplayName="Purchase Date" Name="PurchaseDate" Format="DateOnly" />' },
      { name: 'Cost', xml: '<Field Type="Currency" DisplayName="Cost" Name="Cost" Decimals="2" LCID="1033" />' },
      { name: 'WarrantyMonths', xml: '<Field Type="Number" DisplayName="Warranty (months)" Name="WarrantyMonths" Min="0" />' },
      { name: 'InService', xml: '<Field Type="Boolean" DisplayName="In Service" Name="InService"><Default>1</Default></Field>' },
    ],
    rows: ({ humans }) => {
      const to = (i) => (humans[i % humans.length] ? { AssignedUserId: humans[i % humans.length].Id } : {});
      const rows = [
        ['ThinkPad X1 Carbon G11', 'AST-4401', 'Laptop', '2025-11-14', 2189.0, 36, true],
        ['Dell U2723QE 27"', 'AST-4402', 'Monitor', '2025-11-14', 619.5, 36, true],
        ['iPhone 15 Pro', 'AST-4403', 'Phone', '2026-01-09', 1449.0, 24, true],
        ['Logitech MX Master 3S', 'AST-4404', 'Peripheral', '2026-01-09', 129.99, 12, true],
        ['MacBook Pro 14 M3', 'AST-4405', 'Laptop', '2026-02-27', 2799.0, 36, true],
        ['PowerEdge R660', 'AST-4406', 'Server', '2025-08-21', 11480.0, 60, true],
        ['Dell P2419H 24"', 'AST-4407', 'Monitor', '2023-05-30', 289.0, 36, false],
      ];
      return rows.map((r, i) => ({
        Title: r[0], AssetTag: r[1], Category: r[2], PurchaseDate: d(r[3]),
        Cost: r[4], WarrantyMonths: r[5], InService: r[6], ...to(i),
      }));
    },
  },

  // -------------------------------------------------------------------------
  {
    title: 'Contracts',
    description: 'Restricted — used to demonstrate unique permissions',
    template: 100,
    fields: [
      { name: 'Counterparty', xml: '<Field Type="Text" DisplayName="Counterparty" Name="Counterparty" MaxLength="100" />' },
      { name: 'ContractValue', xml: '<Field Type="Currency" DisplayName="Value" Name="ContractValue" Decimals="2" LCID="1033" />' },
      { name: 'RenewalDate', xml: '<Field Type="DateTime" DisplayName="Renewal Date" Name="RenewalDate" Format="DateOnly" />' },
      { name: 'ContractState', xml: CHOICE('ContractState', 'State', ['Draft', 'Under Review', 'Signed', 'Expired'], 'Draft') },
      { name: 'Confidential', xml: '<Field Type="Boolean" DisplayName="Confidential" Name="Confidential"><Default>1</Default></Field>' },
    ],
    rows: [
      { Title: 'MSA — Northwind Traders', Counterparty: 'Northwind Traders', ContractValue: 240000, RenewalDate: d('2027-03-31'), ContractState: 'Signed', Confidential: true },
      { Title: 'SOW — Fabrikam Migration', Counterparty: 'Fabrikam Inc.', ContractValue: 96500, RenewalDate: d('2026-09-30'), ContractState: 'Under Review', Confidential: true },
      { Title: 'NDA — Tailspin Toys', Counterparty: 'Tailspin Toys', ContractValue: 0, RenewalDate: d('2028-01-15'), ContractState: 'Signed', Confidential: false },
      { Title: 'Support Retainer — Contoso Ltd', Counterparty: 'Contoso Ltd', ContractValue: 54000, RenewalDate: d('2026-12-31'), ContractState: 'Draft', Confidential: true },
    ],
  },
];

const ANNOUNCEMENTS = [
  { Title: 'Intranet redesign enters UAT', Body: 'The new hub navigation is available in UAT. Please log defects in the Issues list with repro steps.' },
  { Title: 'Retention policy review delayed', Body: 'Legal sign-off on the finance retention schedule has slipped; PRJ-1003 stays blocked this sprint.' },
  { Title: 'New asset tagging process', Body: 'All loaned hardware must now carry an AST- tag recorded in the Equipment list before it leaves the office.' },
];

const LIBRARY = {
  title: 'Project Documents',
  description: 'Deliverables, specs and meeting notes by project',
  folders: [
    {
      name: 'PRJ-1001 Intranet Redesign',
      files: [
        { name: 'information-architecture.md', body: '# Information Architecture\n\n## Hub structure\n- Home\n- Departments\n- Policies\n- Projects\n\n## Open questions\n- Does Finance need a separate hub?\n- Retention on archived news posts?\n' },
        { name: 'component-inventory.csv', body: 'component,status,owner,notes\nhero,done,design,uses theme colours\ncard-grid,done,design,\nquick-links,in progress,build,needs icon set\npeople-web-part,blocked,build,waiting on Graph permissions\nevent-list,not started,,\n' },
        { name: 'kickoff-notes.txt', body: 'Kickoff — 2026-02-02\n\nAttendees: PM, design lead, two devs, comms\n\nDecisions:\n- Communication site template, not team site\n- Component library first, pages second\n- UAT gate at 60% completion\n\nActions:\n- Audit existing templates (done)\n- Draft IA (in review)\n' },
      ],
    },
    {
      name: 'PRJ-1002 Document Migration',
      files: [
        { name: 'metadata-crosswalk.csv', body: 'source_field,target_column,type,transform\nDOC_TITLE,Title,Text,trim\nDOC_AUTHOR,Author,User,resolve by email\nDOC_DATE,Created,DateTime,parse dd/mm/yyyy\nDEPT_CODE,Department,Choice,lookup table\nRETENTION,RetentionLabel,Text,map via policy sheet\n' },
        { name: 'migration-runbook.md', body: '# Migration Runbook\n\n1. Freeze the source share (read-only)\n2. Run the crosswalk validation pass\n3. Migrate in department order, smallest first\n4. Spot-check 5% of items per batch\n5. Publish the completion report\n\n## Rollback\nRestore the share from the pre-freeze snapshot and revoke the destination library.\n' },
        { name: 'batch-01-report.csv', body: 'batch,items,succeeded,failed,duration_min\n1,4820,4811,9,64\n2,5140,5140,0,71\n3,3990,3902,88,58\n' },
      ],
    },
    {
      name: 'Policies',
      files: [
        { name: 'external-sharing-policy.md', body: '# External Sharing Policy\n\nGuest access is permitted only on sites tagged `extranet`.\n\n- Links expire after 30 days\n- Anyone-links are disabled tenant-wide\n- Quarterly access review, owned by the site owner\n' },
        { name: 'naming-conventions.md', body: '# Naming Conventions\n\n`<dept>-<purpose>-<env>`\n\nExamples:\n- `fin-reporting-prod`\n- `hr-onboarding-uat`\n\nAvoid spaces, ampersands and apostrophes in site URLs.\n' },
      ],
    },
    {
      name: 'Meeting Notes',
      files: [
        { name: '2026-06-11-steering.txt', body: 'Steering committee — 2026-06-11\n\n- PRJ-1003 remains blocked (legal)\n- Budget reforecast approved for PRJ-1007\n- Accessibility audit moved to Q4\n' },
        { name: '2026-07-02-architecture.txt', body: 'Architecture review — 2026-07-02\n\n- Agreed: no subsites; use hub associations\n- Managed metadata for department, not choice columns\n- Revisit search schema after migration batch 3\n' },
      ],
    },
  ],
};

module.exports = { LISTS, LIBRARY, GROUPS, ANNOUNCEMENTS, PERSONAS };
