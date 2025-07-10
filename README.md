# SharePoint Link Checker - Edge Sidebar Extension

A minimal Edge sidebar extension that checks links on SharePoint sites using Fluent UI design.

## Features

- **SharePoint Site Detection**: Automatically detects if the current page is a SharePoint site
- **Minimal Fluent UI Design**: Clean, modern interface following Microsoft's design system
- **Link Checking**: Analyzes links on SharePoint pages and categorizes them
- **Sidebar Integration**: Operates as an Edge sidebar extension for easy access

## How it Works

1. **Site Detection**: The extension checks if you're on a SharePoint site by looking for:
   - SharePoint domain patterns (sharepoint.com, sharepoint-df.com)
   - SharePoint-specific meta tags and scripts
   - SharePoint page context information

2. **Link Analysis**: When on a SharePoint site, you can click "Check Links" to:
   - Find all links on the current page
   - Categorize them as SharePoint links, external links, email links, etc.
   - Display results with visual indicators

3. **Status Display**: 
   - ✅ Green checkmark for valid SharePoint and email/phone links
   - ⚠️ Yellow warning for external links
   - ❌ Red X for invalid or problematic links

## Installation

1. Open Microsoft Edge
2. Go to `edge://extensions/`
3. Enable "Developer mode" in the left sidebar
4. Click "Load unpacked"
5. Select the extension folder
6. The extension will appear in your extensions list

## Usage

1. Navigate to any website
2. Click the extension icon in the toolbar to open the sidebar
3. If you're on a SharePoint site, you'll see site information and a "Check Links" button
4. If you're not on a SharePoint site, you'll see a message indicating this
5. Click "Check Links" to analyze all links on the SharePoint page

## File Structure

```
SharePoint Link Checker/
├── manifest.json          # Extension manifest
├── sidebar.html           # Sidebar UI
├── sidebar.js             # Sidebar functionality
├── styles.css             # Fluent UI styling
├── background.js          # Background service worker
├── content.js             # Content script for page interaction
├── icons/                 # Extension icons
│   ├── icon16.png
│   ├── icon48.png
│   └── icon128.png
└── README.md              # This file
```

## Technologies Used

- **Manifest V3**: Latest Chrome extension manifest version
- **Fluent UI**: Microsoft's design system for consistent look and feel
- **Edge Sidebar API**: Native sidebar integration
- **Content Scripts**: For page interaction and link analysis

## Permissions

- `activeTab`: To access the current tab's content
- `sidePanel`: To display the sidebar
- SharePoint host permissions for site detection

## Development

The extension is built with vanilla JavaScript and uses:
- Fluent UI CSS for styling
- Chrome Extension APIs for functionality
- Content scripts for SharePoint page interaction

## Notes

- The extension only activates on SharePoint sites
- Link checking is performed client-side for privacy
- External link validation is limited due to CORS restrictions
- Designed specifically for Microsoft Edge browser
