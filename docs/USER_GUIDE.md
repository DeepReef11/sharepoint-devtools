# SharePoint DevTools - User Guide

## Table of Contents
1. [Getting Started](#getting-started)
2. [Installation](#installation)
3. [Features Overview](#features-overview)
4. [Keyboard Shortcuts](#keyboard-shortcuts)
5. [Using QuickNav](#using-quicknav)
6. [Using Column Inspector](#using-column-inspector)
7. [Using Object Inspector](#using-object-inspector)
8. [Managing Favorites and Recent Links](#managing-favorites-and-recent-links)
9. [Custom Links](#custom-links)
10. [Tips and Tricks](#tips-and-tricks)
11. [Troubleshooting](#troubleshooting)

## Getting Started

SharePoint DevTools is a browser extension that enhances your SharePoint Online experience with keyboard-driven navigation and developer tools. It works with both Modern and Classic SharePoint UI.

### Key Benefits
- **Faster Navigation**: Jump to any SharePoint page with keyboard shortcuts
- **Developer Tools**: Inspect column metadata, GUIDs, and internal names
- **Fuzzy Search**: Find what you need even with typos
- **Context-Aware**: Shows relevant links based on your current location
- **Multi-Tenant Support**: Works across different SharePoint environments

## Installation

### Chrome/Edge Installation
1. Download the extension package or build from source
2. Open your browser and navigate to `chrome://extensions/` (Chrome) or `edge://extensions/` (Edge)
3. Enable "Developer mode" (toggle in top-right corner)
4. Click "Load unpacked"
5. Select the `dist` folder from the extension files
6. The SharePoint DevTools icon should appear in your browser toolbar

### Verification
1. Navigate to any SharePoint site (e.g., `https://yourtenant.sharepoint.com`)
2. Press **Ctrl+K** (or **Cmd+K** on Mac) or **Alt+N**
3. The QuickNav modal should appear

## Features Overview

### 🚀 Quick Navigation (Ctrl+K)
Access SharePoint settings, admin pages, and common locations instantly with a searchable modal.

**What you can access:**
- Site settings and permissions
- List/library settings
- Content type management
- Site collection administration
- Recycle bin
- Search settings
- And 75+ more locations

### 📋 Column Inspector (Alt+C)
View detailed metadata about SharePoint list columns, including:
- Internal column names
- Column types
- GUIDs
- Required/Hidden status
- Default values

### 🔍 Object Inspector (Alt+I)
Inspect SharePoint objects and their properties:
- Site collections
- Webs/subsites
- Lists and libraries
- Fields
- Content types

### 💾 Favorites & Recent Links
- Automatically tracks your recently accessed locations
- Mark frequently used links as favorites
- Quick access from the QuickNav modal

### ➕ Custom Links
Add your own SharePoint URLs to the quick navigation:
- Site-specific pages
- Custom admin panels
- Frequently used reports

## Keyboard Shortcuts

| Shortcut | Feature | Description |
|----------|---------|-------------|
| **Ctrl+K** (Cmd+K) or **Alt+N** | QuickNav | Open quick navigation modal |
| **Ctrl+L** (Cmd+L) or **Alt+O** | Lists & Libraries | Browse all lists and libraries in current site |
| **Alt+C** | Column Inspector | Show column metadata for current list |
| **Alt+I** | Object Inspector | Inspect SharePoint objects |
| **↑ / ↓** | Navigate Results | Move up/down in search results |
| **Home / End** | Jump | Jump to first/last result |
| **Enter** | Select | Navigate to selected item |
| **Esc** | Close | Close any open modal |

## Using QuickNav

### Opening QuickNav
Press **Ctrl+K** (or **Cmd+K** on Mac) or **Alt+N** while on any SharePoint page.

### Searching for Links
1. Type in the search box to filter available links
2. The search is fuzzy, so typos are okay (e.g., "setings" will find "Site Settings")
3. Use arrow keys (↑/↓) to navigate results
4. Press **Enter** to navigate to the selected link
5. Press **Esc** to close the modal

### Search Tips
- **Exact matches** appear first
- **Keywords** are weighted heavily (try "permissions", "recycle", "admin")
- **Descriptions** are also searched
- **Category** names are searchable

### Context-Aware Filtering
QuickNav automatically adjusts based on your location:
- **On a site**: Shows site-level actions
- **On a list**: Shows list-specific settings
- **On a library**: Shows library management options

### Organizing Links
- ⭐ **Favorites**: Click the star icon next to any link
- 🕒 **Recent**: Your last 10 accessed links appear at the top
- 🔖 **Custom**: Add your own links via the extension options

## Using Column Inspector

### Opening Column Inspector
1. Navigate to any SharePoint list or library
2. Press **Alt+C**
3. The Column Inspector panel appears on the right side

### What You'll See
Each column displays:
- **Display Name**: The user-friendly name
- **Internal Name**: The system name (useful for APIs and PowerShell)
- **Type**: Field type (Text, Number, Choice, etc.)
- **GUID**: Unique identifier
- **Required**: Whether the field is mandatory
- **Hidden**: Whether the field is hidden from forms
- **Default Value**: Any default value configured

### Copying Values
Click the copy button (📋) next to any value to copy it to your clipboard.

### Use Cases
- **PowerShell/PnP Scripts**: Get internal names for scripting
- **REST API Development**: Get field GUIDs and internal names
- **Troubleshooting**: Check field configurations
- **Documentation**: Export field metadata

## Using Object Inspector

### Opening Object Inspector
1. Navigate to any SharePoint site
2. Press **Alt+I**
3. Select the object type to inspect:
   - Site Collection
   - Web (current site/subsite)
   - Lists/Libraries
   - Fields
   - Content Types

### Inspecting Objects
1. Choose an object type from the dropdown
2. Select a specific object (if applicable)
3. View all properties and metadata
4. Copy GUIDs, URLs, and other values as needed

### Use Cases
- **Getting Site Collection ID**: For app registrations
- **Finding Web URL**: For REST API calls
- **Listing All Lists**: See all lists and libraries with their GUIDs
- **Exploring Content Types**: View content type hierarchy

## Managing Favorites and Recent Links

### Adding Favorites
1. Open QuickNav (Ctrl+K)
2. Find the link you want to favorite
3. Click the ⭐ star icon next to the link
4. The link now appears in your Favorites section

### Removing Favorites
1. Open QuickNav
2. Find the favorited link (marked with ⭐)
3. Click the ⭐ icon again to remove it

### Recent Links
- QuickNav automatically tracks your last 10 accessed links
- Recent links appear at the top of search results
- Clear recent history from the extension options page

## Custom Links

### Adding Custom Links
1. Right-click the extension icon and select "Options"
2. Navigate to the "Custom Links" section
3. Click "Add Custom Link"
4. Fill in:
   - **Title**: Display name for the link
   - **URL**: SharePoint URL (can use placeholders)
   - **Category**: Organizational category
   - **Keywords**: Search keywords (optional)
5. Click "Save"

### URL Placeholders
Use these placeholders in custom link URLs:
- `{siteUrl}`: Current site collection URL
- `{webUrl}`: Current web URL
- `{listId}`: Current list GUID
- `{listUrl}`: Current list URL

**Example:**
```
Title: My Custom Report
URL: {siteUrl}/_layouts/15/CustomReport.aspx?list={listId}
Category: Custom
```

### Editing/Deleting Custom Links
1. Go to extension options
2. Find the custom link
3. Click "Edit" or "Delete"

## Tips and Tricks

### Keyboard Shortcuts Speed
Get faster with these tips:
- Memorize your top 5 most-used shortcuts
- Use fuzzy search - type "perms" instead of "permissions"
- Create custom links for frequently accessed pages

### Search Techniques
- **Partial matching**: "recy" finds "Recycle Bin"
- **Keywords**: Use descriptive terms like "admin", "security", "content"
- **Category filtering**: Include category names in search (e.g., "admin permissions")

### Power User Features
- **Quick GUID Copying**: Alt+C on a list, click copy button
- **Batch Operations**: Use Object Inspector to get multiple list GUIDs at once
- **Custom Workflow**: Create custom links for your specific workflow

### Multi-Tenant Usage
If you work across multiple SharePoint environments (Commercial, GCC, GCC High):
- The extension automatically detects tenant types
- Custom links work across tenants using placeholders
- Favorites are stored per-tenant

## Troubleshooting

### QuickNav Won't Open
**Symptoms**: Pressing Ctrl+K does nothing
**Solutions**:
1. Verify you're on a SharePoint site (URL contains `.sharepoint.com` or `.sharepoint-df.com`)
2. Refresh the page (F5)
3. Check if the extension is enabled in `chrome://extensions/`
4. Try the alternate shortcut: Alt+N

### Column Inspector Shows "No Columns Found"
**Symptoms**: Column Inspector opens but shows no data
**Solutions**:
1. Ensure you're on a list or library page (not a site home page)
2. Wait for the page to fully load before opening
3. Check browser console for errors (F12)
4. Verify you have read permissions on the list

### Search Results Are Empty
**Symptoms**: Search in QuickNav returns no results
**Solutions**:
1. Clear your search and try a shorter query
2. Try searching by category (e.g., "admin", "content")
3. Check if custom links filter is enabled
4. Verify the link exists for your current context (some links only appear on lists)

### Links Don't Navigate
**Symptoms**: Clicking a link doesn't navigate or shows an error
**Solutions**:
1. Verify you have permissions for the target page
2. Check if you're in the correct context (e.g., list links require a list)
3. Some links require specific SharePoint features to be enabled
4. Try opening the link in a new tab (middle-click)

### Extension Not Working After Update
**Solutions**:
1. Disable and re-enable the extension
2. Reload the extension
3. Clear browser cache
4. Restart your browser

### Performance Issues
**Symptoms**: Extension is slow or unresponsive
**Solutions**:
1. Clear recent links history
2. Reduce number of custom links
3. Close other tabs/extensions
4. Check for browser updates

## Feedback and Support

### Reporting Issues
If you encounter bugs or have feature requests:
1. Open an issue on GitHub: [github.com/DeepReef11/sharepoint-devtools/issues](https://github.com/DeepReef11/sharepoint-devtools/issues)
2. Include:
   - Browser version
   - SharePoint environment (Commercial/GCC/GCC High)
   - Steps to reproduce
   - Screenshots if applicable

### Feature Requests
We welcome suggestions for new features! Submit them via GitHub Issues with the "enhancement" label.

## Privacy and Permissions

### Data Storage
- Favorites and recent links are stored locally in your browser
- Custom links are stored in Chrome/Edge sync storage (synced across your devices if browser sync is enabled)
- No data is sent to external servers

### Required Permissions
- **Active Tab**: To detect SharePoint context
- **Storage**: To save favorites and custom links
- **Host Permissions**: SharePoint sites only (*.sharepoint.com, *.sharepoint-df.com)

## Appendix: Full Link Categories

The extension includes links organized into these categories:

1. **Site Admin**: Site settings, permissions, features
2. **Site Collection Admin**: Site collection settings, galleries
3. **List/Library Settings**: List configuration, views, columns
4. **Content Management**: Content types, site columns, metadata
5. **Search**: Search settings and configuration
6. **User Management**: People, groups, permissions
7. **Customization**: Themes, master pages, composed looks
8. **Apps**: App catalog, app settings
9. **Recycle**: Recycle bins (site and site collection)
10. **Reports**: Usage reports, audit logs
11. **Developer**: API explorer, term store, managed metadata

---

**Version**: 1.0.0
**License**: MIT
