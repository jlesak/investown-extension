# Investown Partner Summary

A Chrome extension that enhances [Investown.cz](https://my.investown.cz) by displaying aggregated investment data per partner (borrower) — directly on property detail and listing pages.

## Features

- **Partner Summary Widget** on property detail pages showing:
  - Total investment across all projects by the same partner
  - Total yields earned from the partner
  - Number of projects invested in vs. total projects
- **Partner Column** on the listing page with per-card badges showing partner name, investment total, and project count
- Stale-while-revalidate caching with 5-minute TTL
- Concurrency-limited API requests (max 3 parallel)
- Automatic re-injection when the SPA re-renders the DOM
- Works seamlessly with Investown's single-page application routing

## Installation

Since this extension is not published on the Chrome Web Store, you need to install it manually as an unpacked extension.

### Prerequisites

- Google Chrome (or any Chromium-based browser like Edge, Brave, etc.)
- An active [Investown.cz](https://my.investown.cz) account

### Steps

1. **Download or clone** this repository:

   ```bash
   git clone https://github.com/jlesak/investown-extension.git
   ```

2. **Open the extensions page** in Chrome:
   - Navigate to `chrome://extensions/`
   - Or go to **Menu > Extensions > Manage Extensions**

3. **Enable Developer Mode** using the toggle in the top-right corner.

4. **Click "Load unpacked"** and select the `investown-extension` folder (the root of this repository).

5. The extension icon should appear in your toolbar. Navigate to [my.investown.cz](https://my.investown.cz) and log in — the partner summary will appear automatically.

## Usage

Once installed, the extension works automatically on `my.investown.cz`:

### Property Detail Page

When you open a property detail page (e.g. `my.investown.cz/property/some-project`), a partner summary widget appears in the right sidebar above the "Investice" section. It shows your total investment, total yields, and project count for that partner.

### Listing Page

On the main listing page (`my.investown.cz/`), a "Partner" column is injected into each project card showing the partner name, your total investment with that partner, and how many of their projects you've invested in.

### Authentication

The extension uses your existing Investown session (JWT token from localStorage). You must be logged in for the extension to work. If you're not logged in, it will display a "Nepřihlášen" (not logged in) message.

## How It Works

The extension is a Manifest V3 content script that:

1. Intercepts SPA navigation events (`pushState`, `replaceState`, `popstate`)
2. Detects the current page type (property detail, listing, or other)
3. Queries the Investown GraphQL API for property and related properties data
4. Computes aggregated partner summaries (total investment, yields, project count)
5. Injects the summary widget or listing badges into the DOM
6. Uses a `MutationObserver` to re-inject elements if React removes them during reconciliation

## Project Structure

```
investown-extension/
├── manifest.json     # Chrome extension manifest (Manifest V3)
├── content.js        # Content script — all logic (API, caching, DOM injection)
├── content.css       # Styles for the widget and listing badges
└── icons/            # Extension icons (16, 48, 128 px)
```

## License

This project is licensed under the [MIT License](LICENSE).
