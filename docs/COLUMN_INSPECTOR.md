# Column Inspector

The Column Inspector is a feature that displays detailed metadata about SharePoint list/library columns, including internal names, types, IDs, and formulas.

## Features

### Automatic Detection
- Automatically detects when you're on a SharePoint list or library view
- Shows a floating button (📋) in the bottom-right corner when on a list view
- Hides when navigating away from list views

### Column Metadata Display
The inspector shows comprehensive information for each column:

- **Display Name** - The user-friendly title
- **Internal Name** - The field's internal identifier (used in code/CAML)
- **Static Name** - The static field reference
- **Type** - Field type (Text, Number, Lookup, Calculated, etc.)
- **ID** - Column GUID
- **Group** - The field group/category
- **Badges** - Visual indicators for:
  - Required fields
  - Hidden fields
  - Read-only fields
  - Indexed fields

### Special Field Types

**Calculated Columns**
- Displays the complete formula

**Lookup Columns**
- Shows the source list ID
- Displays the lookup field name

**Choice/Multi-Choice Columns**
- Lists all available choices

**Text Columns**
- Shows maximum length

### List Information
Shows metadata about the current list/library:
- Display title
- Internal name
- List ID (GUID)
- Template type and number
- Server relative URL
- Total column count

## Usage

### Opening the Inspector

**Method 1: Floating Button**
- Click the 📋 button in the bottom-right corner (appears on list views)

**Method 2: Keyboard Shortcut**
- Press `Ctrl+Shift+I` (Windows/Linux)
- Press `Cmd+Shift+I` (Mac)

### Filtering Columns

1. **Search Filter**
   - Type in the search box to filter by:
     - Column title
     - Internal name
     - Field type

2. **Show Hidden Toggle**
   - Check "Show hidden" to display system/hidden columns
   - Uncheck to see only visible columns

### Copying Values

Click the "Copy" button next to any value to copy it to clipboard:
- List internal name
- List ID
- Column internal names
- Column IDs

### Refreshing Data

Click the 🔄 button to reload the column schema from SharePoint.

## UI Elements

### Inspector Panel
- Fixed panel on the right side of the screen (500px wide)
- Dark mode support (automatically adapts to system preference)
- Scrollable content area

### Column Cards
Each column is displayed in a card showing:
- Title with badges
- Internal name (copyable)
- Type
- ID (copyable)
- Additional properties based on field type

### List Info Section
Displays key list metadata at the top of the inspector.

## Technical Details

### API Calls
The inspector uses SharePoint REST API:
```
/_api/web/lists(guid'{listId}')?$expand=Fields,ContentTypes
```

### Detection Logic
Detects list views by checking for:
- `/Lists/` in URL
- `/Forms/AllItems.aspx`
- `/Shared%20Documents/`
- `viewid=` parameter
- `RootFolder=` parameter
- SharePoint page context (`_spPageContextInfo`)

### Performance
- Data is fetched only when the inspector is opened
- Results are cached until refresh is clicked
- Minimal impact on page load time

## Supported SharePoint Versions

- SharePoint Online (Modern UI)
- SharePoint Online (Classic UI)
- SharePoint 2019/2016 (with REST API enabled)

## Column Types Supported

- Text (single line)
- Note (multi-line text)
- Number
- Currency
- DateTime
- Choice
- Multi-Choice
- Lookup
- Boolean (Yes/No)
- User/Person
- Calculated
- Computed
- URL
- Integer
- Counter (ID field)
- GUID
- Taxonomy (Managed Metadata)
- Location
- Geolocation
- Image
- Thumbnail

## Limitations

1. **Permissions Required**
   - User must have read access to the list
   - Some system fields may not be accessible

2. **Modern Pages**
   - Detection works best on list view pages
   - May not work on custom SPFx web parts

3. **Cross-Origin**
   - Must be on the same SharePoint tenant
   - Extension permissions must be granted

## Troubleshooting

### Inspector doesn't appear
- Verify you're on a list/library view page
- Check browser console for errors
- Try refreshing the page

### "Could not detect list ID" error
- The page may not be a standard list view
- Try navigating to "All Items" view
- Check if `_spPageContextInfo` is available

### "Error fetching list schema" error
- You may not have permissions to the list
- SharePoint REST API may be blocked
- Check network tab for failed requests

### Columns not showing
- Check the "Show hidden" checkbox
- Clear the search filter
- Try clicking the refresh button

## Files

- **Types**: `src/types/column-schema.ts`
- **API**: `src/api/column-fetcher.ts`
- **UI**: `src/ui/column-inspector.ts`
- **Integration**: `src/content/column-inspector-integration.ts`

## Future Enhancements

Potential improvements for future versions:
- Export column schema to JSON/CSV
- Compare schemas between lists
- Show column usage statistics
- Edit column properties (if permissions allow)
- View column dependencies
- Show workflows using each column
