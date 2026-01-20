# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run dev          # Start dev server on https://localhost:3000
npm run build        # Production build to dist/
npm run build:dev    # Development build
npm run lint         # Lint src/**/*.ts
npm test             # Run Jest tests
npm run validate     # Validate manifest.xml
npm run clean        # Remove dist/
```

First-time setup:
```bash
npm install
npx office-addin-dev-certs install   # Generate HTTPS certs
```

**Important:** Update the version in `src/taskpane/taskpane.ts` whenever making changes:
```typescript
const VERSION = '1.0.21';  // Increment this
```
This version displays in the debug panel and helps diagnose cache issues.

## Deployment

Uses **GitHub Pages** with **GitHub Actions** for automated deployment.

- Push to `main` branch triggers the deployment workflow
- Build output goes to GitHub Pages
- Production URL: https://ericb-bissell.github.io/rst-word-addin/

To deploy: simply push to `main` and the GitHub Action will build and deploy automatically.

## Project Overview

Microsoft Word Add-in that converts Word documents to reStructuredText (RST). Runs entirely client-side in Word Online/Desktop webview - the server only hosts static files.

**Current version:** 1.0.21
**Deployed:** https://ericb-bissell.github.io/rst-word-addin/

## Architecture

```
src/
├── taskpane/           # UI panel shown in Word sidebar
│   ├── taskpane.ts     # Main UI logic, debug panel, copy functions
│   ├── taskpane.html   # Collapsible debug panel markup
│   └── taskpane.css    # Styles including debug panel
├── commands/           # Ribbon button handlers
├── converter/          # Core conversion pipeline
│   ├── word-to-rst.ts  # Main orchestrator
│   ├── html-parser.ts  # HTML parsing, layout table detection, list parsing
│   ├── rst-formatter.ts # RST output formatting, list indentation
│   ├── types.ts        # TypeScript interfaces (ListItem with indentLevel, etc.)
│   └── directives/     # RST directive generators
│       ├── image.ts    # .. image::
│       ├── figure.ts   # .. figure:: (images with captions)
│       ├── table.ts    # .. table::
│       ├── contents.ts # .. contents::
│       └── custom.ts   # rst_* style → custom directive handler
├── export/             # ZIP export (RST + images folder)
└── utils/
    └── caption-parser.ts  # Parse "Figure 1:", "Table 1:" captions
```

### Processing Pipeline

1. **Word Document** → Office.js `body.getHtml()`
2. **HTML Extraction** → `parseWordHtml()` in html-parser.ts
3. **Element Parsing** → DocumentElement types
4. **RST Formatting** → `formatDocument()` in rst-formatter.ts
5. **Image Extraction** → ZIP creation via jszip
6. **Output** → RST text + ZIP file

### Key Pattern: Word.run() Context

All Word operations use the async context pattern:

```typescript
await Word.run(async (context) => {
  const body = context.document.body;
  const html = body.getHtml();
  await context.sync();
  // Now html.value contains the HTML
});
```

### Key Pattern: Image Extraction via OOXML

Images must be extracted from OOXML, not from HTML blob URLs or `inlinePictures` API.

**Why OOXML?** The `body.inlinePictures` API only returns inline pictures - it misses images inside text boxes, shapes, and floating objects. OOXML contains ALL embedded images.

```typescript
// Get OOXML which contains all images
const ooxmlResult = body.getOoxml();
await context.sync();

// Parse OOXML to extract images from pkg:part elements
function extractImagesFromOoxml(ooxml: string): OoxmlImage[] {
  const parser = new DOMParser();
  const doc = parser.parseFromString(ooxml, 'application/xml');
  const parts = doc.getElementsByTagName('pkg:part');

  const images: OoxmlImage[] = [];
  for (const part of parts) {
    const name = part.getAttribute('pkg:name') || '';
    const contentType = part.getAttribute('pkg:contentType') || '';

    // Images are in /word/media/ folder
    if (name.includes('/media/') && contentType.startsWith('image/')) {
      const binaryData = part.getElementsByTagName('pkg:binaryData')[0];
      if (binaryData?.textContent) {
        images.push({
          name,
          base64: binaryData.textContent.replace(/\s/g, ''),
          contentType,
        });
      }
    }
  }
  return images;
}
```

Blob URLs in Word's HTML (`blob:https://...` or `~WRS{...}`) are not accessible from the add-in's iframe context.

## Critical Implementation Details

### Layout Tables vs Content Tables

Word uses tables for both content and image layout. Detection:

- **Content tables:** `class="MsoTableGrid"` → render as RST tables
- **Layout tables:** `cellpadding="0"` and contain images → extract images only

See `isLayoutTable()` in html-parser.ts.

### Word List Parsing

- `MsoListParagraph` class indicates list items
- `margin-left` style indicates nesting level (0.5in increments)
- List type detected from bullet character (•, o, § vs 1, a, i)

### RST List Indentation

Uses **3-space indent** for nested lists (RST requirement for enumerated lists).

### Heading Hierarchy

```typescript
const styleMap = {
  'Heading1': { underline: '=', overline: true },
  'Heading2': { underline: '=', overline: false },
  'Heading3': { underline: '-', overline: false },
  'Heading4': { underline: '~', overline: false },
};
```

### Custom Directives (rst_* Styles)

Word styles named `rst_<directive>` convert to RST directives:

- `rst_note` → `.. note::`
- `rst_warning` → `.. warning::`
- `rst_code-block` → `.. code-block::`

Content format for arguments/options:
```
[argument]
:option1: value1
:option2: value2
Body content here
```

## Working Features

| Feature | Notes |
|---------|-------|
| Headers H1-H4 | Proper RST underlines with H1 overline |
| Bold/Italic | `**bold**`, `*italic*`, `***both***` |
| Nested lists | 3+ levels, 3-space indent |
| Tables | Grid format with alignment |
| Table captions | `.. table::` directive with `:name:` |
| Images | Extracts from layout tables, `.. image::` |
| Figures | Images with captions use `.. figure::` |
| Copy to clipboard | Preserves indentation |
| Export as ZIP | RST file + images/ folder |
| Debug panel | Collapsible, separate copy button |

## Bug Fixes

*No known bugs - all fixed*

### Recently Fixed

| Bug | Fix |
|-----|-----|
| Blob URL images not exported | Use OOXML extraction via `body.getOoxml()` to get all embedded images. OOXML contains images as base64 in `pkg:part` elements. |
| Ribbon Copy button does nothing | Added fallback: tries clipboard API first, downloads as `document-rst.txt` if clipboard fails in hidden context. |

## Known Limitations

| Limitation | Reason |
|------------|--------|
| **Shapes (Insert → Shapes) cannot be exported** | Word Shapes are stored as DrawingML vector graphics, not raster images. They appear in HTML export but aren't in OOXML media folder. |
| SmartArt graphics | Similar to Shapes - stored as DrawingML, not exportable images |
| Charts | Stored as chart objects, not images |

**Workaround for Shapes:** Convert the Shape to an image in Word first (right-click → "Save as Picture", then Insert → Pictures).

## Feature Plan

### High Priority

| Feature | Description | Implementation Notes |
|---------|-------------|----------------------|
| Footnotes | `[1]_`, `[#]_`, `[#name]_` syntax | Detect Word footnotes via Office.js footnotes API |
| Citations | `[citation]_` references | Detect Word endnotes or bibliography entries |
| Definition lists | Term + indented definition pairs | Detect from Word formatting (bold term + indented para) |
| Field lists | `:field: value` syntax | Detect from Word tab-separated or table patterns |
| Transitions | Horizontal rules (`----`) | Detect Word horizontal lines or page breaks |

### Medium Priority

| Feature | Description | Implementation Notes |
|---------|-------------|----------------------|
| Inline math | `:math:\`expression\`` role | Detect Word equation objects (OMML) |
| Math directive | `.. math::` block equations | Detect Word equation blocks |
| csv-table | `.. csv-table::` directive | Auto-detect simple Word tables as CSV |
| list-table | `.. list-table::` directive | Alternative table format for simple grids |
| Keyboard role | `:kbd:\`Ctrl+C\`` | Detect specific font/style patterns |
| File role | `:file:\`path\`` | Detect monospace paths or custom style |
| GUI roles | `:guilabel:`, `:menuselection:` | Detect quoted UI text or custom styles |

### Low Priority

| Feature | Description | Implementation Notes |
|---------|-------------|----------------------|
| Include directive | `.. include:: file` | Would require multi-file export |
| Raw directive | `.. raw:: html` | Limited Word use case |
| Doctest blocks | `>>> ` Python prompts | Detect from code formatting |
| Option lists | `-a`, `--flag` syntax | Complex parsing from Word |
| Substitutions | `|name|` references | Detect Word field codes or bookmarks |
| Document metadata | Author, date, revision | Extract from Word document properties |

### Won't Implement

| Feature | Reason |
|---------|--------|
| Strikethrough | No standard RST equivalent |
| Underline | No standard RST equivalent (use italic) |
| Text color | RST is plain text only |
| Highlighting | RST is plain text only |

## Testing

Test in **Word Online** across browsers (Edge, Chrome, Firefox).

Sideload for testing:
1. Open Word Online document
2. Home → Add-ins → More Add-ins → Upload My Add-in
3. Select `manifest.xml` (or `manifest.dev.xml` for localhost)

## Technology Stack

- TypeScript + Webpack
- Office.js (Word JavaScript API)
- jszip for ZIP export
- Jest for unit tests
- No backend - static hosting only (GitHub Pages)
