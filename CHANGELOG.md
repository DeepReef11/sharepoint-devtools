# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0]

First public release.

### Navigation

- QuickNav modal (`Ctrl+K` / `Alt+N`) with fuzzy search over a registry of 75+ built-in
  SharePoint destinations
- Lists & Libraries navigator (`Ctrl+L` / `Alt+O`) for browsing every list and library in the
  current site
- Context-aware filtering — list-scoped destinations appear only inside a list or library
- Link templates with `{webUrl}`, `{siteUrl}`, `{listId}`, `{listUrl}` and `{tenantAdminUrl}`
  placeholders
- Recent and favorite link tracking, plus user-defined custom links managed from the options page
- Multi-tenant support across the Commercial, GCC, GCC High and China clouds

### Inspectors

- Column Inspector (`Alt+C`) — internal names, field types, GUIDs, calculated-column formulas,
  lookup relationships, and the current item's value for each column
- Object Inspector (`Alt+I`) — site, web, list, field and content-type metadata
- Permission Inspector — role assignments, SharePoint groups with expandable membership, and
  permission-level definitions
- Copy-to-clipboard for GUIDs, internal names and field values throughout

### Interface

- Full keyboard control: `↑`/`↓`, `PgUp`/`PgDn`, `Home`/`End`, `Enter`, `Esc`
- Two-level keyboard navigation in the inspectors — move between cards, then into a card's actions
- Automatic dark mode following the system colour scheme
- Support for both modern and classic SharePoint interfaces
- Asynchronous loading so the UI stays responsive while the REST API is queried

### Security

- All SharePoint-supplied values (list titles, column names, item values, group descriptions)
  are HTML-escaped before rendering, including inside HTML attributes
- Rendered links are restricted to `http`, `https` and `mailto`, so a `javascript:` or `data:`
  payload stored in a URL column cannot execute
