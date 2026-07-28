# Modal UI Component

A keyboard-driven modal component for SharePoint DevTools extension.

## Features

- **Clean Design**: Modern, minimal interface that integrates seamlessly with SharePoint
- **Search Input**: Real-time search with debouncing
- **Results List**: Displays results with title and description
- **Keyboard Navigation**: Full keyboard support for efficient navigation
  - `↑`/`↓` - Navigate through results
  - `Enter` - Select current result
  - `Esc` - Close modal
  - `Home` - Jump to first result
  - `End` - Jump to last result
- **Category Grouping**: Results organized by categories
- **Dark Mode Support**: Automatically adapts to system color scheme
- **Loading & Error States**: Visual feedback for different states
- **Smooth Animations**: Fade-in/out and scale animations

## Usage

### Basic Usage

```typescript
import { Modal } from './ui/modal';

const modal = new Modal({
  title: 'SharePoint DevTools',
  placeholder: 'Search...',
  darkMode: false,
  onSearch: (query) => {
    console.log('Search:', query);
  },
  onSelect: (item) => {
    console.log('Selected:', item);
  },
  onClose: () => {
    console.log('Modal closed');
  },
});

// Show modal
modal.show();

// Close modal
modal.close();

// Toggle modal
modal.toggle();
```

### Setting Results

```typescript
// Flat list of results
const results = [
  {
    id: '1',
    title: 'Site Settings',
    description: 'Access site settings',
    url: '/settings',
    category: 'Admin',
  },
];

modal.setResults(results);
```

### Setting Grouped Results

```typescript
// Grouped by category
const groups = [
  {
    category: 'Site Admin',
    items: [
      {
        id: '1',
        title: 'Site Settings',
        description: 'Access site settings',
        url: '/settings',
        category: 'Site Admin',
      },
    ],
  },
  {
    category: 'Content',
    items: [
      {
        id: '2',
        title: 'Document Libraries',
        description: 'View all libraries',
        url: '/libraries',
        category: 'Content',
      },
    ],
  },
];

modal.setGroupedResults(groups);
```

### Setting State

```typescript
// Show loading state
modal.setState('loading', 'Searching...');

// Show error state
modal.setState('error', 'An error occurred');

// Return to idle state
modal.setState('idle');
```

### Dark Mode

```typescript
// Enable dark mode
modal.setDarkMode(true);

// Disable dark mode
modal.setDarkMode(false);
```

## API

### Constructor Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `title` | `string` | `'SharePoint DevTools'` | Modal title |
| `placeholder` | `string` | `'Search for SharePoint locations...'` | Search input placeholder |
| `darkMode` | `boolean` | `false` | Enable dark mode |
| `onSearch` | `(query: string) => void` | - | Search callback (debounced 300ms) |
| `onSelect` | `(item: ResultItem) => void` | - | Item selection callback |
| `onClose` | `() => void` | - | Modal close callback |

### Methods

| Method | Description |
|--------|-------------|
| `show()` | Show the modal |
| `close()` | Close the modal |
| `toggle()` | Toggle modal visibility |
| `setResults(results: ResultItem[])` | Set flat list of results |
| `setGroupedResults(groups: CategoryGroup[])` | Set grouped results |
| `setState(state: ModalState, message?: string)` | Set modal state |
| `setDarkMode(enabled: boolean)` | Toggle dark mode |
| `isOpen()` | Check if modal is visible |
| `destroy()` | Destroy modal and clean up |

### Types

```typescript
interface ResultItem {
  id: string;
  title: string;
  description?: string;
  url?: string;
  category: string;
  metadata?: Record<string, any>;
}

interface CategoryGroup {
  category: string;
  items: ResultItem[];
}

type ModalState = 'idle' | 'loading' | 'error' | 'success';
```

## Styling

The modal uses CSS with the `spqn-` prefix to avoid conflicts. All styles are defined in `modal.css` and support both light and dark modes.

### CSS Classes

- `.spqn-overlay` - Modal overlay
- `.spqn-modal` - Modal container
- `.spqn-search-input` - Search input
- `.spqn-results-container` - Results list container
- `.spqn-result-item` - Individual result item
- `.spqn-category-header` - Category header
- `.spqn-dark-mode` - Dark mode modifier

## Browser Compatibility

- Chrome/Edge (Manifest V3)
- Modern browsers with CSS Grid and Flexbox support
- Requires ES6+ JavaScript support
