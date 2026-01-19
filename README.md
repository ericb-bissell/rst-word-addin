# RST Word Add-in

A Microsoft Word Add-in for previewing and exporting Word documents as reStructuredText (RST).

## Features

- **Live Preview** - See your Word document as RST in real-time
- **Copy to Clipboard** - Copy RST text for use in other applications
- **Export as ZIP** - Download RST file with extracted images
- **Custom Directives** - Use `rst_*` Word styles for Sphinx/RST directives
- **Full RST Support** - Headings, lists, tables, images, figures, and more

## Quick Start

### Prerequisites

- Node.js 18 or higher
- Microsoft 365 account (for Word Online)

### Installation

```bash
# Install dependencies
npm install

# Install HTTPS certificates (first time only)
npx office-addin-dev-certs install

# Start development server
npm run dev
```

### Sideload the Add-in

1. Open [Word Online](https://office.com)
2. Open or create a document
3. Go to **Home** → **Add-ins** → **More Add-ins**
4. Click **Upload My Add-in**
5. Select `manifest.xml` from the `dist/` folder
6. Click the **RST Preview** button in the ribbon

## Development

```bash
npm run dev      # Start dev server with hot reload
npm run build    # Build for production
npm run lint     # Run linter
npm test         # Run tests
```

## Project Structure

```
office_rst/
├── src/
│   ├── taskpane/        # Taskpane UI
│   ├── commands/        # Ribbon commands
│   ├── converter/       # HTML to RST conversion
│   ├── export/          # ZIP export functionality
│   └── utils/           # Shared utilities
├── assets/              # Icons and images
├── docs/                # User documentation
├── dist/                # Built files (generated)
└── manifest.xml         # Office Add-in manifest
```

## Documentation

- [User Guide](docs/README.md) - How to use the add-in
- [Formatting Reference](docs/FORMATTING.md) - RST formatting support
- [Custom Styles Guide](docs/CUSTOM_STYLES.md) - Creating `rst_*` directive styles
- [Development Guide](Claude.md) - Technical documentation for developers

## Custom RST Directives

Create Word styles starting with `rst_` to generate RST directives:

| Word Style | RST Output |
|------------|------------|
| `rst_note` | `.. note::` |
| `rst_warning` | `.. warning::` |
| `rst_code-block` | `.. code-block::` |

See the [Custom Styles Guide](docs/CUSTOM_STYLES.md) for details.

## License

MIT
