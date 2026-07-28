# Testing the SharePoint DevTools Extension

This guide will help you test the extension in your browser.

## Prerequisites

- Chrome, Edge, or any Chromium-based browser
- Access to a SharePoint site for testing

## Installation Steps

### 1. Build the Extension

```bash
npm install
npm run build
```

The extension will be built to the `dist/` folder.

**Expected output**: `webpack 5.103.0 compiled successfully`

The `dist/` folder should contain:
- `background.js` - Background service worker
- `content.js` - Content script
- `options.js` / `options.html` - Options page
- `manifest.json` - Extension manifest
- `icons/` - Extension icons (16px, 48px, 128px)

### 2. Load Extension in Chrome/Edge

1. Open your browser and navigate to:
   - **Chrome**: `chrome://extensions`
   - **Edge**: `edge://extensions`

2. Enable **Developer mode** (toggle in top-right corner)

3. Click **"Load unpacked"**

4. Select the `dist/` folder from this project

5. The extension should now appear in your extensions list

### 3. Verify Installation

You should see:
- Extension icon in your browser toolbar
- Extension enabled with no errors

## Testing Features

### Feature 1: QuickNav Modal (Alt+N)

**What it does**: Quick navigation to SharePoint locations and settings

**How to test**:
1. Navigate to any SharePoint site (e.g., `https://yourtenant.sharepoint.com`)
2. Either:
   - Press **Alt+N** keyboard shortcut, OR
   - Click the extension icon and click "Open QuickNav"
3. A modal should appear with "SharePoint DevTools" title
4. Try typing to search for locations
5. Use **↑/↓** arrows to navigate results
6. Press **Enter** to select or **Esc** to close

**Expected result**: Modal opens and allows searching/navigation

---

### Feature 2: Lists & Libraries Navigator (Alt+O)

**What it does**: Browse and search all lists and libraries in the current SharePoint site

**How to test**:
1. Navigate to any SharePoint site
2. Press **Alt+O** keyboard shortcut
3. A modal titled "Lists & Libraries" should appear
4. You should see all non-hidden lists and libraries from the current site
5. Each item shows:
   - List/Library name
   - Type (List or Library)
   - Item count
6. Try typing to search/filter
7. Use **↑/↓** arrows to navigate results
8. Press **Enter** to navigate to the selected list/library
9. Press **Esc** to close

**Expected result**: Modal displays all site lists/libraries with fast filtering

---

### Feature 3: Column Inspector (Alt+C)

**What it does**: Displays internal column names in SharePoint lists

**How to test**:
1. Navigate to a SharePoint list or library (e.g., Documents, Custom List)
2. Press **Alt+C**
3. A side panel should appear showing column metadata
4. You should see:
   - Column display names
   - Internal names
   - Field types
   - Copy buttons for easy copying

**Expected result**: Side panel with column information appears

---

### Feature 3: Object Inspector (Alt+I)

**What it does**: Inspects SharePoint objects (sites, webs, lists, items)

**How to test**:
1. Navigate to any SharePoint page
2. Press **Alt+I**
3. A modal titled "SharePoint Object Inspector" should appear
4. You should see:
   - Current context (site/web/list)
   - Metadata and properties
   - REST API responses
   - Copy-to-clipboard buttons

**Expected result**: Inspector modal shows SharePoint object metadata

---

### Feature 4: Fuzzy Search

**What it does**: Smart search that handles typos

**How to test**:
1. Open QuickNav modal (**Alt+S**)
2. Type a search with intentional typos:
   - "setings" (should find "Site Settings")
   - "permisions" (should find "Site Permissions")
   - "recy" (should find "Recycle Bin")
3. Results should still appear despite typos

**Expected result**: Fuzzy matching finds results even with typos

---

### Feature 5: Custom Links (Options Page)

**What it does**: Manage custom SharePoint links

**How to test**:
1. Right-click extension icon → **Options**
2. The options page should open
3. Try:
   - Adding a custom link
   - Editing a link
   - Deleting a link
   - Import/export links

**Expected result**: Full CRUD operations work

---

### Feature 6: Recent & Favorite Links

**What it does**: Tracks recently visited links and favorites

**How to test**:
1. Open QuickNav (**Alt+S**)
2. Navigate to several SharePoint locations
3. Re-open QuickNav
4. You should see recently visited links
5. Try marking links as favorites

**Expected result**: Recent links appear, favorites can be saved

---

### Feature 7: Multi-Tenant Support

**What it does**: Works across multiple SharePoint tenants

**How to test**:
1. Test on different SharePoint environments:
   - Production: `tenant.sharepoint.com`
   - GCC: `tenant.sharepoint.us`
   - DoD: `tenant.sharepoint-df.com`
2. Extension should auto-detect and work on all

**Expected result**: Works on all SharePoint variants

---

## Keyboard Shortcuts Summary

| Shortcut | Feature | Description |
|----------|---------|-------------|
| **Alt+N** | QuickNav | Open quick navigation modal |
| **Alt+O** | Lists & Libraries | Browse all lists and libraries |
| **Alt+C** | Column Inspector | Show column metadata panel |
| **Alt+I** | Object Inspector | Inspect SharePoint objects |
| **↑/↓** | Navigate | Navigate search results |
| **Enter** | Select | Navigate to selected item |
| **Esc** | Close | Close any modal |

### Alternative Access
- **Click extension icon** → Opens popup with quick access buttons
- **Right-click icon → Options** → Opens settings page

## Debugging

### Check Console Logs

1. On any SharePoint page, open DevTools (**F12**)
2. Go to **Console** tab
3. Look for messages like:
   - `SharePoint DevTools - Content Script loaded`
   - `SharePoint page detected`
   - `QuickNav modal initialized`
   - `Column Inspector initialized`
   - `Object Inspector initialized`

### Common Issues

**Issue**: Extension doesn't load
- **Fix**: Check Extensions page for errors
- **Fix**: Reload the extension
- **Fix**: Rebuild with `npm run build`

**Issue**: Keyboard shortcuts don't work
- **Fix**: Make sure you're on a SharePoint page (URL contains `.sharepoint.com`)
- **Fix**: Check no other extension is using the same shortcut

**Issue**: No data appears in inspectors
- **Fix**: Check browser console for API errors
- **Fix**: Verify you have permissions to the SharePoint site
- **Fix**: Check network tab for REST API calls

## Development Mode

For active development:

```bash
# Watch mode - rebuilds on file changes
npm run watch

# Development build with source maps
npm run dev
```

After code changes:
1. Save files
2. Go to `chrome://extensions`
3. Click reload icon on the extension
4. Refresh the SharePoint page

## Production Build

For production use:

```bash
npm run build
```

This creates an optimized build in `dist/`.

## Test Checklist

- [ ] Extension installs without errors
- [ ] QuickNav modal opens (Alt+N)
- [ ] Search works in QuickNav
- [ ] Lists & Libraries modal opens (Alt+O)
- [ ] Lists modal shows all site lists/libraries
- [ ] Search/filter works in Lists modal
- [ ] Keyboard navigation works (↑/↓/Enter/Esc)
- [ ] Column Inspector shows metadata (Alt+C)
- [ ] Object Inspector opens (Alt+I)
- [ ] Fuzzy search handles typos
- [ ] Custom links can be added/edited
- [ ] Recent links are tracked
- [ ] Works on different SharePoint sites
- [ ] No console errors

## Need Help?

- Check console logs in DevTools (F12)
- Review `ROADMAP.md` for feature status
- Open an issue on GitHub with:
  - Browser version
  - SharePoint environment
  - Console errors
  - Steps to reproduce
