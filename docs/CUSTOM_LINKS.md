# Custom Links Management

**Feature Status**: Implemented
**Version**: 1.0.0

## Overview

The Custom Links Management feature allows users to create, manage, and organize their own SharePoint navigation links alongside the built-in links provided by SharePoint DevTools. This feature gives users the flexibility to add shortcuts to frequently accessed pages, custom applications, or any SharePoint resource they need quick access to.

## Features

### Core Functionality

- **Create Custom Links**: Add new navigation links with custom titles, URLs, categories, and descriptions
- **Edit Links**: Modify existing custom links at any time
- **Delete Links**: Remove custom links that are no longer needed
- **Import/Export**: Share custom link collections or back them up as JSON files
- **Link Validation**: Test URL templates with sample placeholder values before saving
- **Storage Management**: View storage usage and manage your custom links collection

### Link Properties

Each custom link supports the following properties:

- **ID** (required): Unique identifier (lowercase, hyphens allowed)
- **Title** (required): Display name for the link
- **URL Template** (required): URL with placeholders like `{webUrl}`, `{siteUrl}`, `{listId}`, `{listUrl}`
- **Category** (required): Organizes links into groups (Site Admin, Content, Navigation, etc.)
- **Description** (required): Explains what the link does
- **Keywords** (optional): Comma-separated search terms for finding the link
- **Priority** (optional): Lower numbers appear first in search results
- **Context Requirements** (optional):
  - Requires List/Library Context
  - Requires Site Admin permissions
  - Requires Tenant Admin permissions
  - Modern SharePoint Only
  - Classic SharePoint Only

## Accessing Custom Links Settings

There are two ways to access the Custom Links Management interface:

1. **Extension Icon**: Click the SharePoint DevTools extension icon in your browser toolbar, then select "Options" or "Settings"
2. **Extension Menu**: Right-click the extension icon and select "Options"

## Usage Guide

### Creating a New Custom Link

1. Open the Custom Links settings page
2. Click the **"Add New Link"** button
3. Fill in the required fields:
   - **Link ID**: A unique identifier (e.g., `my-custom-page`)
   - **Title**: Display name (e.g., "My Custom Dashboard")
   - **URL Template**: SharePoint URL with placeholders
   - **Category**: Choose from available categories
   - **Description**: Brief explanation of what the link does
4. Optionally add:
   - Keywords for easier searching
   - Priority value for result ordering
   - Context requirements for conditional visibility
5. Click **"Test Link"** to verify the URL resolves correctly
6. Click **"Save Link"** to add the link to your collection

### Example Custom Links

#### Custom Application Page
```json
{
  "id": "custom-dashboard",
  "title": "Team Dashboard",
  "urlTemplate": "{webUrl}/SitePages/TeamDashboard.aspx",
  "category": "Content",
  "description": "Custom team performance dashboard",
  "keywords": ["dashboard", "team", "metrics"]
}
```

#### Custom List Settings
```json
{
  "id": "project-tracker-settings",
  "title": "Project Tracker Settings",
  "urlTemplate": "{webUrl}/_layouts/15/listedit.aspx?List={listId}",
  "category": "List/Library Settings",
  "description": "Settings for the Project Tracker list",
  "context": {
    "requiresList": true
  }
}
```

#### Custom Admin Tool
```json
{
  "id": "custom-admin-tool",
  "title": "Custom Site Provisioning Tool",
  "urlTemplate": "{siteUrl}/_layouts/15/CustomProvisioning.aspx",
  "category": "Development",
  "description": "Internal site provisioning tool",
  "context": {
    "requiresSiteAdmin": true,
    "modernOnly": true
  },
  "priority": 5
}
```

### Editing Custom Links

1. Find the link you want to edit in the list
2. Click the **"Edit"** button on the link card
3. Modify the fields as needed
4. Click **"Save Link"** to apply changes

### Deleting Custom Links

1. Find the link you want to delete
2. Click the **"Delete"** button on the link card
3. Confirm the deletion in the popup dialog

### Importing Custom Links

You can import custom links from a JSON file:

1. Click the **"Import"** button
2. Choose import mode:
   - **Replace**: Delete all existing custom links and import new ones
   - **Merge**: Keep existing links and add/update with imported ones
3. Either:
   - Click **"Choose JSON file"** to select a file, OR
   - Paste JSON data directly into the text area
4. Click **"Import Links"** to complete the import

### Exporting Custom Links

To back up or share your custom links:

1. Click the **"Export"** button
2. A JSON file will be downloaded automatically
3. The filename will be `sharepoint-devtools-links-YYYY-MM-DD.json`

### Clearing All Custom Links

To remove all custom links at once:

1. Click the **"Clear All"** button
2. Confirm the action in the popup dialog
3. ⚠️ **Warning**: This action cannot be undone. Consider exporting first!

## URL Placeholders

Custom links support the following placeholder values that are automatically replaced based on the current SharePoint context:

| Placeholder | Description | Example |
|-------------|-------------|---------|
| `{webUrl}` | Current web/subsite URL | `https://contoso.sharepoint.com/sites/mysite` |
| `{siteUrl}` | Root site collection URL | `https://contoso.sharepoint.com/sites/mysite` |
| `{listId}` | Current list/library GUID | `{12345678-1234-1234-1234-123456789012}` |
| `{listUrl}` | Current list/library URL | `https://contoso.sharepoint.com/sites/mysite/Lists/MyList` |

### Placeholder Usage Tips

- **List Context**: Links using `{listId}` or `{listUrl}` require the user to be viewing a list or library
- **Multiple Placeholders**: You can use multiple placeholders in a single URL
- **Testing**: Use the "Test Link" button to see how your URL resolves with sample values
- **Context Requirements**: Set appropriate context requirements to ensure placeholders are available

## Integration with QuickNav

Custom links are seamlessly integrated with the main QuickNav interface:

1. **Automatic Loading**: Custom links are loaded when the extension initializes
2. **Search Integration**: Custom links appear in search results alongside built-in links
3. **Category Grouping**: Custom links are grouped by category just like built-in links
4. **Priority Sorting**: Links with priority values appear first in search results
5. **Context Filtering**: Links are automatically filtered based on context requirements

## Storage and Sync

- **Storage Backend**: Chrome Extension `chrome.storage.sync`
- **Cross-Device Sync**: Custom links automatically sync across devices when you're signed into Chrome
- **Storage Limit**: Chrome sync storage has a quota of ~100KB (can store hundreds of links)
- **Version Management**: Link data includes version information for future migrations

## Best Practices

### Naming Conventions

- **IDs**: Use lowercase with hyphens (e.g., `my-custom-link`)
- **Titles**: Use clear, descriptive names (e.g., "Project Status Dashboard")
- **Keywords**: Include common search terms and synonyms

### Organization

- **Categories**: Choose categories that match your workflow
- **Priorities**: Use priorities sparingly for your most important links
- **Descriptions**: Write clear descriptions to help future you remember what the link does

### URL Templates

- **Test First**: Always test your URL templates before saving
- **Validate Placeholders**: Ensure placeholders match your context requirements
- **Document Special Cases**: Add notes in descriptions for non-obvious URL patterns

### Backup and Sharing

- **Regular Exports**: Export your links periodically as a backup
- **Team Sharing**: Share exported JSON files with team members
- **Version Control**: Consider storing exported JSON in a shared repository

## Troubleshooting

### Link Not Appearing in QuickNav

- Verify the link has the correct category
- Check if context requirements are preventing the link from showing
- Ensure the URL template can be resolved in the current context
- Try reloading the extension

### Import Fails

- Validate JSON syntax (use a JSON validator)
- Ensure all required fields are present
- Check that IDs are unique and properly formatted
- Verify the JSON structure matches the expected format

### Storage Quota Exceeded

- Export and remove unnecessary links
- Consider splitting links across multiple browser profiles
- Check storage usage in the settings page statistics

### Link URL Not Resolving

- Verify placeholder spelling and casing
- Check that required placeholders are available in current context
- Test the URL template using the "Test Link" button
- Ensure context requirements match the placeholder usage

## Technical Details

### File Structure

```
src/
├── storage/
│   └── custom-links-storage.ts    # Storage management module
├── options/
│   ├── options.html                # Settings page UI
│   ├── options.css                 # Styles
│   └── options.ts                  # Settings page controller
└── links/
    └── link-manager.ts             # Updated to support custom links
```

### API Reference

See the TypeScript interfaces in `src/storage/custom-links-storage.ts` for detailed API documentation.

Key classes:
- `CustomLinksStorage`: Handles all storage operations
- `OptionsController`: Manages the settings page UI

### Data Format

Custom links are stored in this JSON format:

```json
{
  "version": "1.0",
  "links": [
    {
      "id": "example-link",
      "title": "Example Link",
      "urlTemplate": "{webUrl}/example",
      "category": "Content",
      "description": "An example custom link",
      "keywords": ["example", "demo"],
      "priority": 10,
      "context": {
        "requiresList": false,
        "modernOnly": true
      }
    }
  ],
  "lastModified": "2026-01-15T09:00:00.000Z"
}
```

## Future Enhancements

Potential improvements for future versions:

- [ ] Link templates/presets for common patterns
- [ ] Bulk operations (multi-select, bulk delete)
- [ ] Link categories management (create custom categories)
- [ ] Usage statistics (track most-used custom links)
- [ ] Advanced filtering and sorting in settings UI
- [ ] Link validation with actual SharePoint API calls
- [ ] Collaborative link sharing via cloud service
- [ ] Import from browser bookmarks

## Related Documentation

- [SharePoint DevTools Overview](../README.md)
- [Column Inspector Documentation](./COLUMN_INSPECTOR.md)
- [Development Roadmap](../ROADMAP.md)

## Support

For issues or questions:
- Check the [Troubleshooting](#troubleshooting) section above
- Review the GitHub Issues for similar problems
- Create a new issue with detailed information about your problem

---

**Feature Version**: 1.0.0
**Roadmap Issue**: #11
