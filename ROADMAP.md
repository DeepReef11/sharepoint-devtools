# Roadmap

Planned work, roughly in priority order. Suggestions are welcome via
[GitHub Issues](https://github.com/DeepReef11/sharepoint-devtools/issues).

## Next

- **Store distribution** — Chrome Web Store and Edge Add-ons listings: promotional assets,
  screenshots and a privacy policy
- **TypeScript strict mode** — `strict` is currently off; enabling it will surface a backlog of
  nullability fixes that are worth paying down
- **Quieter logging** — several modules log to the console unconditionally; these should be behind
  a debug flag rather than shipping in production builds

## Under consideration

- **Content type inspector** — inspect content types and their field bindings directly, rather
  than linking out to the settings page
- **View inspector** — surface view CAML, row limits and filters for the current list view
- **Search-driven navigation** — resolve a site or list by name from anywhere in the tenant,
  not only within the current site
- **Export** — copy a list's full column schema as JSON or as a PnP PowerShell snippet
- **Firefox support** — the extension targets Manifest V3; Firefox compatibility needs
  investigation

## Non-goals

- Editing SharePoint data. The inspectors are read-only by design; the extension links out to
  SharePoint's own settings pages for anything that mutates state.
- Bundling a UI framework. The extension is intentionally dependency-light so the content script
  stays small on every SharePoint page load.
