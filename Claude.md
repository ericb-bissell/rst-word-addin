# RST Word Add-in - Project Instructions

## Current Status (v1.0.13)

**Deployed:** GitHub Pages at `https://ericb-bissell.github.io/rst-word-addin/`
**Repo:** `ericb-bissell/rst-word-addin`

### Working Features

| Feature | Status | Notes |
|---------|--------|-------|
| Headers (H1-H4) | ✅ Working | Proper RST underlines (=, -, ~, ^) with H1 overline |
| Bold/Italic | ✅ Working | `**bold**`, `*italic*`, `***both***` |
| Bullet Lists | ✅ Working | Nested to 3+ levels with 3-space indent |
| Numbered Lists | ✅ Working | Auto-numbering with `#.`, nested support |
| Tables | ✅ Working | Grid format with proper cell alignment |
| Table Captions | ✅ Working | `.. table::` directive with `:name:` |
| Images | ✅ Working | Extracts from layout tables, `.. image::` directive |
| Image Attributes | ✅ Working | alt, width, height preserved |
| Copy to Clipboard | ✅ Working | Clean RST only, preserves indentation |
| Export as ZIP | ✅ Working | RST file + images folder |
| Debug Panel | ✅ Working | Collapsible, separate copy button |

### Recent Version History

| Version | Changes |
|---------|---------|
| 1.0.13 | Collapsible debug panel, copy preserves indentation |
| 1.0.12 | Layout table image detection (cellpadding=0 tables) |
| 1.0.11 | Fixed RST list indentation (3-space for nested) |
| 1.0.10 | Nested list support with indent level detection |
| 1.0.9 | Table caption association fix |

### Key Implementation Files

| File | Purpose |
|------|---------|
| `src/converter/html-parser.ts` | HTML parsing, layout table detection, list parsing |
| `src/converter/rst-formatter.ts` | RST output formatting, list indentation |
| `src/converter/types.ts` | TypeScript interfaces including ListItem with indentLevel |
| `src/taskpane/taskpane.ts` | UI logic, debug panel, copy functions |
| `src/taskpane/taskpane.html` | Collapsible debug panel markup |
| `src/taskpane/taskpane.css` | Debug panel styles |

### Known Patterns

**Layout Tables vs Content Tables:**
- Content tables have `class="MsoTableGrid"`
- Layout tables have `cellpadding="0"` and contain images
- `isLayoutTable()` function distinguishes them

**Word List Parsing:**
- `MsoListParagraph` class indicates list items
- `margin-left` style indicates nesting level (0.5in increments)
- List type detected from bullet character (•, o, § vs 1, a, i)

**RST List Indentation:**
- Uses 3-space indent for nested lists (RST requirement for enumerated)
- Auto-numbering uses `#.` marker

---

## Project Overview

This is a Microsoft Office 365 Add-in for Word that enables users to preview and export Word documents as reStructuredText (RST) format. The add-in provides:

- **Live RST Preview**: View the Word document content as RST in a side panel
- **Copy to Clipboard**: Copy the RST content for use in other applications
- **Export as RST**: Save the document as a ZIP containing `.rst` file and extracted images

Similar to the [Markdown Word Add-in](https://markdownword.barinbritva.com/), but for reStructuredText.

### Key Features

| Feature | Description |
|---------|-------------|
| **Image Export** | Extracts images to `images/` folder, referenced via `.. image::` directive with full attribute support |
| **Figure Support** | Images with captions use `.. figure::` directive preserving caption text and numbering |
| **Table Directives** | Tables with captions use `.. table::` directive with alignment and width attributes |
| **Table of Contents** | Word TOC fields convert to `.. contents::` directive with depth control |
| **Full RST Attributes** | All supported directive options (alt, width, height, scale, align, etc.) are preserved |
| **Custom Directives** | Word styles named `rst_*` convert to custom RST directives (e.g., `rst_need` → `.. need::`) for Sphinx extensions |

## Testing Environment

This add-in is developed and tested using **Word Online (Web version)** of Office 365.

### Sideloading for Testing

1. Open [Office on the web](https://office.com/) and open a Word document
2. Select **Home** > **Add-ins** > **More Add-ins**
3. Select **Upload My Add-in** (or **My Add-ins** > **Upload My Add-in**)
4. Browse to the `manifest.xml` file and upload it
5. The add-in will appear in the ribbon under the configured tab

### Development Server

```bash
npm run dev       # Start development server
npm run start     # Start with Word Online (requires document URL)
npm run build     # Build for production
```

## Architecture

### Runtime Model (Client-Side)

Office Add-ins run **entirely client-side** in the user's browser or Office desktop webview. There is no backend server processing - the server only hosts static files.

```
┌─────────────────────────────────────────────────────────────┐
│  User's Browser / Office Desktop Webview                    │
│  ┌────────────────────────────────────────────────────────┐ │
│  │  Word Application                                      │ │
│  │  ┌──────────────────────────────────────────────────┐  │ │
│  │  │  Add-in iframe                                   │  │ │
│  │  │  • HTML/CSS/JS runs HERE (client-side)           │  │ │
│  │  │  • Office.js API calls happen HERE               │  │ │
│  │  │  • RST conversion happens HERE                   │  │ │
│  │  │  • All processing is LOCAL to user's machine     │  │ │
│  │  └──────────────────────────────────────────────────┘  │ │
│  └────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘
                            │
                            │ Initial load only (fetches static files once)
                            ▼
                   ┌──────────────────┐
                   │  Static Server   │
                   │  (serves files)  │
                   │  No processing   │
                   └──────────────────┘
```

**Key implications:**

| Aspect | Reality |
|--------|---------|
| Code execution | 100% client-side in user's browser |
| Server role | Serves static files only (like a CDN) |
| Backend required | No - all conversion logic runs in JavaScript |
| Development | Local HTTPS server (`https://localhost:3000`) |
| Production hosting | Any static host: GitHub Pages, Azure Static Web Apps, Netlify, S3 |
| Cost | Minimal - static hosting is cheap/free |

This is why we implement the RST converter in TypeScript/JavaScript rather than using Python's `docutils` - everything must run in the browser.

### Project Structure

```
office_rst/
├── manifest.xml                 # Office Add-in manifest (XML format)
├── package.json                 # Node.js dependencies
├── webpack.config.js            # Webpack bundler configuration
├── tsconfig.json                # TypeScript configuration
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html        # Task pane UI
│   │   ├── taskpane.css         # Task pane styles
│   │   ├── taskpane.ts          # Task pane logic (Office.js API)
│   │   ├── help.html            # In-app help page
│   │   ├── help.css             # Help page styles
│   │   └── help.js              # Help page navigation logic
│   ├── commands/
│   │   └── commands.ts          # Ribbon button command handlers
│   ├── converter/
│   │   ├── word-to-rst.ts       # Main Word document to RST conversion
│   │   ├── html-parser.ts       # Parse Word HTML output
│   │   ├── rst-formatter.ts     # RST syntax formatting utilities
│   │   └── directives/
│   │       ├── image.ts         # .. image:: directive generator
│   │       ├── figure.ts        # .. figure:: directive generator
│   │       ├── table.ts         # .. table:: directive generator
│   │       ├── contents.ts      # .. contents:: directive generator
│   │       └── custom.ts        # rst_* style → custom directive handler
│   ├── export/
│   │   ├── image-extractor.ts   # Extract images from Word document
│   │   ├── zip-builder.ts       # Bundle RST + images into ZIP
│   │   └── file-download.ts     # Browser file download helpers
│   └── utils/
│       ├── office-helpers.ts    # Office.js helper functions
│       └── caption-parser.ts    # Parse Word captions (Figure 1:, Table 1:)
├── assets/
│   ├── icon-16.png              # Add-in icons (16x16)
│   ├── icon-32.png              # Add-in icons (32x32)
│   ├── icon-80.png              # Add-in icons (80x80)
│   └── logo-filled.png          # Logo for task pane
├── dist/                        # Built files (generated)
├── docs/                        # User documentation
│   ├── README.md                # Main user guide
│   ├── FORMATTING.md            # RST formatting reference
│   └── CUSTOM_STYLES.md         # Custom directive styles guide
└── tests/
    ├── converter.test.ts        # Unit tests for main converter
    ├── directives.test.ts       # Unit tests for directive generators
    └── fixtures/                # Test HTML/document fixtures
        └── sample-doc.html
```

### Technology Stack

- **Framework**: TypeScript + Webpack
- **Office API**: Word JavaScript API (Office.js)
- **UI**: HTML/CSS (Fluent UI optional)
- **RST Conversion**: Custom converter (no reliable JS library exists for HTML→RST)
- **Testing**: Jest for unit tests, manual testing in Word Online

## Word JavaScript API - Key Concepts

### Basic Pattern

All Word operations use the `Word.run()` context pattern:

```typescript
await Word.run(async (context) => {
  // 1. Access document objects
  const body = context.document.body;

  // 2. Queue property loads
  body.load("text");
  const paragraphs = body.paragraphs;
  paragraphs.load("items");

  // 3. Execute queued commands
  await context.sync();

  // 4. Work with loaded data
  console.log(body.text);
  paragraphs.items.forEach(p => console.log(p.text));
});
```

### Getting Document Content

```typescript
// Get HTML representation (useful for conversion)
const html = body.getHtml();
await context.sync();
console.log(html.value);

// Get plain text
body.load("text");
await context.sync();
console.log(body.text);

// Get OOXML (full fidelity)
const ooxml = body.getOoxml();
await context.sync();
console.log(ooxml.value);
```

### Key Collections

- `body.paragraphs` - All paragraphs
- `body.tables` - All tables
- `body.inlinePictures` - Inline images
- `body.lists` - List objects
- `body.contentControls` - Content controls

### Paragraph Properties

```typescript
paragraphs.load("text, style, styleBuiltIn, font, alignment, firstLineIndent");
await context.sync();

paragraphs.items.forEach(p => {
  console.log(`Text: ${p.text}`);
  console.log(`Style: ${p.styleBuiltIn}`);  // e.g., "Heading1", "Normal"
  console.log(`Bold: ${p.font.bold}`);
});
```

## RST Conversion Strategy

### Approach: HTML → RST

Use `body.getHtml()` to get HTML representation, then convert to RST. This is more reliable than parsing OOXML.

### Export Output Structure

When exporting, the add-in creates a structured output:

```
export/
├── document.rst              # Main RST document
└── images/                   # Extracted images folder
    ├── image_001.png
    ├── image_002.jpg
    └── figure_003.png
```

Images are extracted from the Word document, saved as separate files, and referenced in the RST using relative paths.

### RST Markup Reference

| Element | RST Syntax |
|---------|------------|
| Heading 1 | Text + `===` underline (with overline) |
| Heading 2 | Text + `===` underline |
| Heading 3 | Text + `---` underline |
| Heading 4 | Text + `~~~` underline |
| Bold | `**text**` |
| Italic | `*text*` |
| Inline code | ``` ``code`` ``` |
| Code block | `::` + indented block |
| Link | `` `text <url>`__ `` |
| Image | `.. image::` directive |
| Figure | `.. figure::` directive (image with caption) |
| Table | `.. table::` directive |
| TOC | `.. contents::` directive |
| Unordered list | `- item` |
| Ordered list | `1. item` |
| Block quote | Indented text |

---

### Image Directive (`.. image::`)

For images **without captions**, use the `image` directive with all applicable options extracted from Word.

**Full Syntax:**

```rst
.. image:: images/screenshot.png
   :alt: Alternative text description
   :height: 200px
   :width: 400px
   :scale: 75%
   :align: center
   :target: https://example.com
   :class: custom-class
   :name: image-reference-name
   :loading: lazy
```

**Supported Options:**

| Option | Type | Description |
|--------|------|-------------|
| `:alt:` | text | Alternate text for accessibility (defaults to filename) |
| `:height:` | length | Desired height (e.g., `200px`, `5cm`) |
| `:width:` | length/% | Width as length or percentage (e.g., `400px`, `80%`) |
| `:scale:` | integer % | Uniform scaling factor (e.g., `75%`) |
| `:align:` | enum | `top`, `middle`, `bottom` (inline) or `left`, `center`, `right` (block) |
| `:target:` | URI | Makes image clickable, links to URI |
| `:class:` | text | Space-separated CSS class names |
| `:name:` | text | Reference name for cross-references |
| `:loading:` | enum | `embed`, `link`, or `lazy` |

**Word → RST Mapping:**

```typescript
interface ImageOptions {
  path: string;           // Relative path to saved image file
  altText?: string;       // From Word alt text
  width?: string;         // From Word image width
  height?: string;        // From Word image height
  alignment?: 'left' | 'center' | 'right';  // From Word alignment
  hyperlink?: string;     // If image is a hyperlink in Word
}
```

---

### Figure Directive (`.. figure::`)

For images **with captions or labels**, use the `figure` directive. This wraps the image in a figure container with caption and optional legend.

**Full Syntax:**

```rst
.. figure:: images/architecture.png
   :alt: System architecture diagram
   :height: 300px
   :width: 600px
   :scale: 100%
   :align: center
   :figwidth: 80%
   :figclass: diagram
   :name: fig-architecture

   This is the figure caption (single paragraph).

   This is the legend, which can contain multiple paragraphs
   and other body elements like lists:

   - Item one
   - Item two
```

**Figure-Specific Options:**

| Option | Type | Description |
|--------|------|-------------|
| `:figwidth:` | length/% or `image` | Width of entire figure container |
| `:figclass:` | text | CSS class names on the figure element |
| `:figname:` | text | Reference name for the figure element |
| `:align:` | enum | `left`, `center`, or `right` (horizontal only) |

**Inherited Image Options:** All `:alt:`, `:height:`, `:width:`, `:scale:`, `:target:`, `:class:`, `:name:`, `:loading:` options.

**Word → RST Mapping:**

```typescript
interface FigureOptions extends ImageOptions {
  caption?: string;       // From Word caption (e.g., "Figure 1: Description")
  figureNumber?: string;  // Extracted figure number
  figWidth?: string;      // Container width
  legend?: string;        // Additional descriptive text
}

// Detection: Use figure directive when image has:
// - A caption paragraph (Word "Insert Caption" feature)
// - A label like "Figure 1:", "Fig. 1:", etc.
```

---

### Table Directive (`.. table::`)

Tables with titles or special formatting use the `table` directive wrapper. Tables without titles can use raw grid/simple table syntax.

**Full Syntax:**

```rst
.. table:: Table 1: Sales Data by Quarter
   :align: center
   :width: 100%
   :widths: 20 30 25 25
   :class: sales-table
   :name: tbl-sales

   +----------+----------+----------+----------+
   | Quarter  | Revenue  | Costs    | Profit   |
   +==========+==========+==========+==========+
   | Q1       | $50,000  | $30,000  | $20,000  |
   +----------+----------+----------+----------+
   | Q2       | $65,000  | $35,000  | $30,000  |
   +----------+----------+----------+----------+
```

**Table Options:**

| Option | Type | Description |
|--------|------|-------------|
| `:align:` | enum | `left`, `center`, or `right` |
| `:width:` | length/% | Total table width |
| `:widths:` | list/auto/grid | Column widths as integers, `auto`, or `grid` |
| `:class:` | text | Space-separated CSS class names |
| `:name:` | text | Reference name for cross-references |

**Word → RST Mapping:**

```typescript
interface TableOptions {
  caption?: string;       // From Word table caption
  tableNumber?: string;   // Extracted table number (e.g., "Table 1")
  alignment?: 'left' | 'center' | 'right';
  width?: string;         // Total width
  columnWidths?: number[]; // Relative column widths
  hasHeader?: boolean;    // First row is header
}

// Detection: Use table directive when table has:
// - A caption (Word "Insert Caption" feature)
// - A label like "Table 1:", "Tbl. 1:", etc.
// - Special formatting attributes to preserve
```

**Table Formats:**

Grid tables (complex, supports spanning):
```rst
+----------+----------+
| Header 1 | Header 2 |
+==========+==========+
| Cell 1   | Cell 2   |
+----------+----------+
```

Simple tables (easier to write):
```rst
========  ========
Header 1  Header 2
========  ========
Cell 1    Cell 2
Cell 3    Cell 4
========  ========
```

---

### Contents Directive (`.. contents::`)

Generates an automatic table of contents from document headings. Detected when Word document contains a TOC field.

**Full Syntax:**

```rst
.. contents:: Table of Contents
   :depth: 3
   :local:
   :backlinks: entry
   :class: toc-custom
```

**Contents Options:**

| Option | Type | Description |
|--------|------|-------------|
| `:depth:` | integer | Max heading levels to include (default: unlimited) |
| `:local:` | flag | Only include subsections of current section |
| `:backlinks:` | enum | `entry` (link to TOC entry), `top` (link to TOC), or `none` |
| `:class:` | text | Space-separated CSS class names |

**Word → RST Mapping:**

```typescript
interface ContentsOptions {
  title?: string;         // TOC title (default: "Contents")
  depth?: number;         // From Word TOC heading level setting
  includePageNumbers?: boolean;  // Note: RST doesn't support this natively
}

// Detection: Look for Word TOC field codes or
// paragraph styles like "TOC Heading", "TOC 1", "TOC 2", etc.
```

---

### Custom Directives (`rst_*` Styles)

Word styles with names beginning with `rst_` are automatically converted to RST directives. This enables support for Sphinx extensions, custom directives, and domain-specific markup.

**Convention:**

| Word Style Name | RST Directive |
|-----------------|---------------|
| `rst_note` | `.. note::` |
| `rst_warning` | `.. warning::` |
| `rst_need` | `.. need::` |
| `rst_req` | `.. req::` |
| `rst_code-block` | `.. code-block::` |
| `rst_admonition` | `.. admonition::` |

**Example - Word Document:**

A paragraph styled with `rst_need` containing:
```
REQ-001: The system shall support user authentication
```

**Converts to RST:**

```rst
.. need::

   REQ-001: The system shall support user authentication
```

**Advanced: Directive Arguments and Options**

For directives that require arguments or options, use a structured format in the Word content:

```
[argument]
:option1: value1
:option2: value2
Content goes here
```

**Example - Code Block with Language:**

Word paragraph with style `rst_code-block` containing:
```
[python]
:linenos:
def hello():
    print("Hello, World!")
```

**Converts to:**

```rst
.. code-block:: python
   :linenos:

   def hello():
       print("Hello, World!")
```

**Example - Sphinx-Needs Requirement:**

Word paragraph with style `rst_need` containing:
```
[req]
:id: REQ-001
:title: User Authentication
:status: open
The system shall support user authentication via OAuth 2.0.
```

**Converts to:**

```rst
.. need:: req
   :id: REQ-001
   :title: User Authentication
   :status: open

   The system shall support user authentication via OAuth 2.0.
```

**Word → RST Mapping:**

```typescript
interface CustomDirective {
  name: string;           // Directive name (extracted from style after "rst_")
  argument?: string;      // Optional argument (from [argument] syntax)
  options: Map<string, string>;  // Key-value options (from :key: value lines)
  content: string;        // Remaining content (indented in output)
}

// Detection: Check paragraph.style or paragraph.styleBuiltIn
// If style starts with "rst_", extract directive name
function parseCustomDirective(styleName: string, content: string): CustomDirective {
  const directiveName = styleName.replace(/^rst_/, '').replace(/-/g, '-');

  // Parse content for [argument], :options:, and body
  const lines = content.split('\n');
  let argument: string | undefined;
  const options = new Map<string, string>();
  const bodyLines: string[] = [];

  for (const line of lines) {
    if (line.startsWith('[') && line.endsWith(']') && !argument) {
      argument = line.slice(1, -1);
    } else if (line.match(/^:\w+:/)) {
      const match = line.match(/^:(\w+):\s*(.*)$/);
      if (match) options.set(match[1], match[2]);
    } else {
      bodyLines.push(line);
    }
  }

  return {
    name: directiveName,
    argument,
    options,
    content: bodyLines.join('\n').trim()
  };
}
```

**Built-in RST Directives Supported via `rst_*` Styles:**

| Category | Directives |
|----------|------------|
| Admonitions | `rst_note`, `rst_warning`, `rst_tip`, `rst_important`, `rst_caution`, `rst_danger`, `rst_error`, `rst_hint`, `rst_attention` |
| Code | `rst_code-block`, `rst_literalinclude`, `rst_code` |
| Structure | `rst_topic`, `rst_sidebar`, `rst_rubric`, `rst_epigraph`, `rst_highlights`, `rst_pull-quote` |
| Sphinx | `rst_toctree`, `rst_only`, `rst_index`, `rst_glossary`, `rst_productionlist` |
| Extensions | `rst_need`, `rst_req`, `rst_spec`, `rst_test` (sphinx-needs), any custom directive |

**Creating Custom Styles in Word:**

Users can create `rst_*` styles in Word to use custom directives:

1. **Home tab** → **Styles** → Click the small arrow to open Styles pane
2. Click **New Style** button at bottom of pane
3. Set **Name** to `rst_` followed by directive name (e.g., `rst_need`, `rst_warning`)
4. Set **Style type** to `Paragraph`
5. Optionally configure formatting to visually distinguish directive content
6. Click **OK** to save

Recommended style formatting for visual distinction:
- **Background color**: Light gray or colored background
- **Left indent**: 0.5" to visually indicate directive content
- **Font**: Monospace font (Consolas, Courier New) for code-related directives
- **Border**: Left border to highlight directive blocks

---

### Heading Hierarchy

RST uses underline characters for headings. Recommended hierarchy:

```rst
=================
Document Title (H1)
=================

Section (H2)
============

Subsection (H3)
---------------

Sub-subsection (H4)
~~~~~~~~~~~~~~~~~~~

Paragraph heading (H5)
^^^^^^^^^^^^^^^^^^^^^^

Minor heading (H6)
""""""""""""""""""
```

### Conversion Mapping

```typescript
// Word Style → RST
const styleMap = {
  'Heading1': { underline: '=', overline: true },
  'Heading2': { underline: '=', overline: false },
  'Heading3': { underline: '-', overline: false },
  'Heading4': { underline: '~', overline: false },
  'Heading5': { underline: '^', overline: false },
  'Heading6': { underline: '"', overline: false },
};
```

---

### Image Extraction and Saving

Images are extracted from the Word document and saved to disk during export:

```typescript
interface ExtractedImage {
  id: string;              // Unique identifier
  filename: string;        // Generated filename (e.g., "image_001.png")
  data: Blob;              // Image binary data
  format: string;          // "png", "jpg", "gif", etc.
  originalName?: string;   // Original filename if available
  width?: number;          // Original width in pixels
  height?: number;         // Original height in pixels
}

// Export process:
// 1. Extract images via Word.js InlinePicture API
// 2. Get base64 data from each image
// 3. Convert to Blob and determine format
// 4. Generate unique filename
// 5. Bundle into ZIP with RST file for download
```

**Export Download Format:**

The export creates a ZIP file containing:
- `document.rst` - The converted RST document
- `images/` - Folder with all extracted images

This ensures all references resolve correctly when the user extracts the ZIP.

## Manifest Configuration

### Key Sections

```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="TaskPaneApp">

  <Id>GUID-HERE</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Your Name</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="RST Preview"/>
  <Description DefaultValue="Preview and export Word documents as reStructuredText"/>

  <Hosts>
    <Host Name="Document"/>  <!-- Word -->
  </Hosts>

  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="..." xsi:type="VersionOverridesV1_0">
    <!-- Ribbon commands defined here -->
  </VersionOverrides>
</OfficeApp>
```

### Ribbon Integration

Add a custom tab or group with buttons:
- **Preview RST**: Opens task pane with live preview
- **Export RST**: Downloads ZIP containing `.rst` file and `images/` folder
- **Copy RST**: Copies RST text to clipboard (without images)

## Development Guidelines

### Error Handling

Always wrap Office.js calls in try-catch:

```typescript
try {
  await Word.run(async (context) => {
    // ... operations
  });
} catch (error) {
  if (error instanceof OfficeExtension.Error) {
    console.error(`Office Error: ${error.code} - ${error.message}`);
  }
  throw error;
}
```

### Performance

- Load only needed properties: `paragraphs.load("text, styleBuiltIn")`
- Batch operations within a single `Word.run()`
- Use `context.sync()` sparingly (batch multiple operations)

### Browser Compatibility

Word Online runs in the browser. Ensure:
- No Node.js-specific APIs in client code
- Use web-compatible file download methods (Blob + download link)
- Test in multiple browsers (Edge, Chrome, Firefox)

## Multi-Agent Development Plan

### Phase 1: Project Setup (Agent: Setup)

**Tasks:**
1. Initialize npm project with TypeScript
2. Install dependencies: `office-addin-dev-certs`, `webpack`, `webpack-dev-server`, `jszip`
3. Configure `tsconfig.json` for Office.js development
4. Configure `webpack.config.js` with dev server and HTTPS
5. Create basic `manifest.xml` with Word host configuration
6. Create placeholder HTML/CSS/TS files
7. Set up project folder structure (converter/, directives/, export/, utils/)

**Deliverables:**
- Working dev server at `https://localhost:3000`
- Valid manifest that can be sideloaded
- Basic "Hello World" task pane

### Phase 2: RST Directives (Agent: Directives)

**Tasks:**
1. Create `directives/image.ts` - Generate `.. image::` with all options (alt, width, height, scale, align, target, class, name, loading)
2. Create `directives/figure.ts` - Generate `.. figure::` with caption, legend, and figure-specific options (figwidth, figclass, figname)
3. Create `directives/table.ts` - Generate `.. table::` with caption, alignment, width, widths, and grid/simple table formatting
4. Create `directives/contents.ts` - Generate `.. contents::` with title, depth, local, backlinks options
5. Create `directives/custom.ts` - Handle `rst_*` Word styles as custom directives:
   - Parse style name to extract directive name (e.g., `rst_need` → `.. need::`)
   - Parse content for `[argument]` syntax
   - Parse content for `:option: value` lines
   - Format remaining content as indented directive body
6. Create `utils/caption-parser.ts` - Parse Word captions ("Figure 1:", "Table 1:") to extract labels and numbers
7. Write unit tests for each directive generator including custom directive parsing

**Deliverables:**
- Complete directive generator modules
- Custom `rst_*` style directive handler with argument/option parsing
- Caption detection and parsing utilities
- Unit tests with sample inputs/outputs

### Phase 3: Core Converter (Agent: Converter)

**Tasks:**
1. Implement `html-parser.ts` for Word's HTML output
2. Create `rst-formatter.ts` with basic markup rules (headings, bold, italic, links, lists)
3. Handle heading hierarchy with proper underline characters
4. Handle text formatting (bold, italic, underline, strikethrough)
5. Handle lists (ordered and unordered, nested)
6. Handle links (inline and reference styles)
7. Handle code blocks and inline code
8. Integrate directive modules for images, figures, tables, TOC
9. Detect TOC fields and convert to `.. contents::` directive
10. Detect `rst_*` paragraph styles and route to custom directive handler
11. Handle consecutive `rst_*` styled paragraphs as multi-paragraph directive content

**Deliverables:**
- `word-to-rst.ts` - Main conversion orchestrator
- `html-parser.ts` - HTML to intermediate representation
- `rst-formatter.ts` - RST syntax utilities
- Full integration with directive modules (including custom `rst_*` directives)

### Phase 4: Image Extraction & Export (Agent: Export)

**Tasks:**
1. Create `export/image-extractor.ts` - Extract images via Word.js InlinePicture API
2. Get image base64 data, detect format (PNG, JPG, GIF, etc.)
3. Generate unique filenames and track image references
4. Create `export/zip-builder.ts` - Bundle RST + images folder using JSZip
5. Create `export/file-download.ts` - Browser-compatible ZIP download
6. Handle image alt text, dimensions, and positioning from Word

**Deliverables:**
- Working image extraction from Word documents
- ZIP file generation with proper folder structure
- Browser download functionality

### Phase 5: Task Pane UI (Agent: UI)

**Tasks:**
1. Design task pane layout (preview area, toolbar)
2. Implement RST preview display with monospace font
3. Add syntax highlighting for RST (optional)
4. Implement copy-to-clipboard button (RST text only)
5. Implement export button (downloads ZIP with RST + images)
6. Add refresh/sync button for live updates
7. Show export progress indicator for large documents
8. Style with Fluent UI or custom CSS matching Office theme

**Deliverables:**
- Responsive task pane UI
- Working copy and export functions
- Clean, Office-consistent styling

### Phase 6: Word Integration (Agent: Integration)

**Tasks:**
1. Implement document content extraction using Word.js API
2. Extract paragraphs, tables, images, and TOC fields
3. Wire up task pane to converter pipeline
4. Implement live preview updates on document change (if feasible)
5. Add ribbon commands for quick actions
6. Handle edge cases (empty documents, unsupported content)

**Deliverables:**
- Full Word.js integration
- Working preview flow
- Ribbon integration

### Phase 7: Polish & Testing (Agent: QA)

**Tasks:**
1. Test in Word Online across browsers (Edge, Chrome, Firefox)
2. Test with various document structures:
   - Documents with many images
   - Documents with captioned figures and tables
   - Documents with nested lists and complex tables
   - Documents with TOC
3. Verify image extraction and ZIP structure
4. Fix edge cases and formatting issues
5. Add user-friendly error messages
6. Optimize performance for large documents
7. Create usage documentation

**Deliverables:**
- Tested, stable add-in
- Known limitations documented
- User guide

## Key Dependencies

```json
{
  "dependencies": {
    "jszip": "^3.10.0"
  },
  "devDependencies": {
    "@types/office-js": "^1.0.0",
    "@types/jszip": "^3.4.0",
    "typescript": "^5.0.0",
    "webpack": "^5.0.0",
    "webpack-cli": "^5.0.0",
    "webpack-dev-server": "^4.0.0",
    "html-webpack-plugin": "^5.0.0",
    "copy-webpack-plugin": "^11.0.0",
    "ts-loader": "^9.0.0",
    "office-addin-dev-certs": "^1.0.0",
    "jest": "^29.0.0",
    "ts-jest": "^29.0.0",
    "@types/jest": "^29.0.0"
  }
}
```

| Package | Purpose |
|---------|---------|
| `jszip` | Create ZIP files containing RST document and extracted images |
| `@types/office-js` | TypeScript definitions for Office.js API |
| `office-addin-dev-certs` | Generate HTTPS certificates for local development |
| `jest` / `ts-jest` | Unit testing for converter and directive modules |
```

## Production Deployment

Since the add-in is entirely client-side, production deployment only requires static file hosting:

### Free Hosting Options

| Service | Notes |
|---------|-------|
| **GitHub Pages** | Free, easy CI/CD with GitHub Actions |
| **Netlify** | Free tier, automatic builds from git |
| **Vercel** | Free tier, great for static sites |
| **Azure Static Web Apps** | Free tier, Microsoft ecosystem |
| **Cloudflare Pages** | Free, fast global CDN |

### Deployment Steps

1. Build production assets: `npm run build`
2. Upload `dist/` folder contents to static host
3. Update `manifest.xml` URLs to point to production host
4. Submit to Microsoft AppSource (optional) or distribute manifest directly

### Manifest Distribution

For internal/personal use, users can sideload the manifest directly - no app store submission required.

## Useful Links

### User Documentation

**In-App Help** (accessible from taskpane Help button):
- `src/taskpane/help.html` - Tabbed help interface with all user documentation

**Repository Documentation** (for GitHub/developers):
- [User Guide](docs/README.md) - Installation, features, and usage
- [Formatting Reference](docs/FORMATTING.md) - How Word formatting converts to RST
- [Custom Styles Guide](docs/CUSTOM_STYLES.md) - Creating `rst_*` directive styles

### External Resources

- [Word Add-ins Overview](https://learn.microsoft.com/en-us/office/dev/add-ins/word/)
- [Word JavaScript API Reference](https://learn.microsoft.com/en-us/javascript/api/word)
- [Office Add-in Manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests)
- [Sideload Add-ins for Testing](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
- [reStructuredText Specification](https://docutils.sourceforge.io/docs/ref/rst/restructuredtext.html)
- [Sphinx-Needs Documentation](https://sphinx-needs.readthedocs.io/)
- [Markdown Word Add-in (Reference)](https://markdownword.barinbritva.com/)

## Known Limitations

1. **Images**: Word Online may not provide full image data; placeholder paths may be needed
2. **Complex Tables**: Nested tables and merged cells have limited RST support
3. **Floating Objects**: Only inline content is reliably accessible via API
4. **Real-time Preview**: Document change events may be limited in Word Online
5. **OOXML**: While available, parsing OOXML is complex; HTML approach is preferred

## Commands Reference

```bash
# Development
npm install              # Install dependencies
npm run dev              # Start dev server
npm run build            # Production build
npm run lint             # Run linter

# Testing
npm test                 # Run unit tests
npm run validate         # Validate manifest

# Certificates (first time setup)
npx office-addin-dev-certs install
```
