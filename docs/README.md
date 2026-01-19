# RST Word Add-in User Guide

Convert your Microsoft Word documents to reStructuredText (RST) format with ease.

## Table of Contents

- [Overview](#overview)
- [Installation](#installation)
- [Getting Started](#getting-started)
- [Features](#features)
- [Using the Add-in](#using-the-add-in)
- [Custom Directives](#custom-directives)
- [Supported Formatting](#supported-formatting)
- [Export Options](#export-options)
- [Troubleshooting](#troubleshooting)

---

## Overview

The RST Word Add-in enables you to:

- **Preview** your Word document as reStructuredText in real-time
- **Copy** the RST content to your clipboard
- **Export** your document as a complete RST package (`.rst` file + images)

This is ideal for technical writers, documentation teams, and anyone who needs to convert Word documents to RST for use with Sphinx, Read the Docs, or other documentation systems.

---

## Installation

### For Word Online (Web Version)

1. Open [Office.com](https://office.com) and sign in
2. Open a Word document (or create a new one)
3. Click **Home** tab in the ribbon
4. Click **Add-ins** → **More Add-ins**
5. Search for "RST Preview" or upload the add-in manifest:
   - Click **My Add-ins** → **Upload My Add-in**
   - Browse to the `manifest.xml` file
   - Click **Upload**
6. The RST add-in will appear in your ribbon

### For Word Desktop (Windows/Mac)

1. Open Word and go to **Insert** → **Get Add-ins**
2. Search for "RST Preview" in the Office Add-ins store
3. Click **Add** to install
4. The add-in will appear in your ribbon

---

## Getting Started

### Quick Start

1. Open any Word document
2. Click the **RST Preview** button in the ribbon
3. A side panel opens showing your document as RST
4. Edit your Word document - the preview updates automatically
5. Click **Copy** to copy RST to clipboard, or **Export** to download

### Your First Export

1. Create a simple Word document with:
   - A title (Heading 1)
   - Some paragraphs
   - A bulleted list
   - An image

2. Click **RST Preview** to open the panel

3. Click **Export** to download a ZIP file containing:
   ```
   your-document.zip
   ├── document.rst      # Your converted document
   └── images/           # Extracted images
       └── image_001.png
   ```

4. Extract the ZIP and use `document.rst` in your Sphinx project

---

## Features

### Live Preview

The preview panel shows your document converted to RST in real-time. As you edit your Word document, the preview updates to reflect your changes.

### Copy to Clipboard

Click **Copy** to copy the RST text to your clipboard. This is useful for:
- Pasting into a text editor
- Quick snippets for documentation
- Sharing via chat or email

Note: Images are not included when copying - use Export for complete documents with images.

### Export as ZIP

Click **Export** to download a complete package:
- `document.rst` - The converted RST file
- `images/` folder - All extracted images with proper references

The RST file uses relative paths (`images/image_001.png`) so everything works when you extract the ZIP.

---

## Using the Add-in

### The Preview Panel

| Button | Action |
|--------|--------|
| **Refresh** | Manually refresh the preview |
| **Copy** | Copy RST text to clipboard |
| **Export** | Download ZIP with RST + images |

### Ribbon Commands

| Command | Description |
|---------|-------------|
| **RST Preview** | Open/close the preview panel |
| **Quick Export** | Export without opening preview |
| **Copy RST** | Copy RST directly to clipboard |

---

## Custom Directives

The add-in supports custom RST directives through special Word styles. Any paragraph style starting with `rst_` is converted to an RST directive.

### Creating Custom Styles

1. Go to **Home** tab → **Styles** panel
2. Click the small arrow (↘) to open the Styles pane
3. Click **New Style** at the bottom
4. Name your style starting with `rst_` (e.g., `rst_note`, `rst_warning`, `rst_need`)
5. Set **Style type** to "Paragraph"
6. Click **OK**

### Using Custom Styles

Apply your `rst_*` style to a paragraph, and it becomes a directive:

**In Word (style: `rst_note`):**
```
Remember to save your work frequently.
```

**RST Output:**
```rst
.. note::

   Remember to save your work frequently.
```

### Advanced: Directives with Arguments and Options

For directives that need arguments or options, format your content like this:

```
[argument]
:option1: value1
:option2: value2
Your content here
```

**Example - Code Block (style: `rst_code-block`):**
```
[python]
:linenos:
def hello():
    print("Hello!")
```

**RST Output:**
```rst
.. code-block:: python
   :linenos:

   def hello():
       print("Hello!")
```

**Example - Sphinx-Needs Requirement (style: `rst_need`):**
```
[req]
:id: REQ-001
:title: User Login
:status: open
The system shall provide a login page.
```

**RST Output:**
```rst
.. need:: req
   :id: REQ-001
   :title: User Login
   :status: open

   The system shall provide a login page.
```

### Recommended Custom Styles

| Style Name | Use For |
|------------|---------|
| `rst_note` | Important notes |
| `rst_warning` | Warning messages |
| `rst_tip` | Helpful tips |
| `rst_code-block` | Code examples with syntax highlighting |
| `rst_need` | Sphinx-needs requirements |
| `rst_todo` | TODO items |

See [CUSTOM_STYLES.md](CUSTOM_STYLES.md) for a complete guide.

---

## Supported Formatting

### Text Formatting

| Word Format | RST Output |
|-------------|------------|
| **Bold** | `**bold**` |
| *Italic* | `*italic*` |
| `Code` (monospace font) | ``` ``code`` ``` |
| ~~Strikethrough~~ | Not supported in standard RST |
| Underline | Not supported in standard RST |

### Headings

Word heading styles convert to RST section titles:

| Word Style | RST Format |
|------------|------------|
| Heading 1 | `====` with overline |
| Heading 2 | `====` underline only |
| Heading 3 | `----` underline |
| Heading 4 | `~~~~` underline |
| Heading 5 | `^^^^` underline |
| Heading 6 | `""""` underline |

### Lists

**Bulleted lists:**
```rst
- First item
- Second item
  - Nested item
- Third item
```

**Numbered lists:**
```rst
1. First item
2. Second item
3. Third item
```

### Links

Hyperlinks in Word convert to RST inline links:

```rst
`Link Text <https://example.com>`__
```

### Images

Images are extracted and referenced using the `.. image::` directive:

```rst
.. image:: images/screenshot.png
   :alt: Screenshot description
   :width: 400px
   :align: center
```

### Images with Captions (Figures)

If your image has a Word caption, it uses the `.. figure::` directive:

```rst
.. figure:: images/architecture.png
   :alt: System architecture
   :width: 600px
   :align: center

   Figure 1: System architecture overview
```

### Tables

Tables convert to RST grid table format:

```rst
.. table:: Table 1: Sales Data
   :align: center

   +----------+----------+----------+
   | Quarter  | Revenue  | Profit   |
   +==========+==========+==========+
   | Q1       | $50,000  | $20,000  |
   +----------+----------+----------+
   | Q2       | $65,000  | $30,000  |
   +----------+----------+----------+
```

### Table of Contents

Word TOC fields convert to the RST contents directive:

```rst
.. contents:: Table of Contents
   :depth: 3
```

See [FORMATTING.md](FORMATTING.md) for a complete formatting reference.

---

## Export Options

### ZIP Package Contents

When you export, you receive a ZIP file with:

```
document-name.zip
├── document.rst          # Main RST document
└── images/               # All images from your document
    ├── image_001.png
    ├── image_002.jpg
    └── figure_003.png
```

### Image Naming

Images are automatically named based on their position:
- `image_001.png` - First image without caption
- `figure_002.png` - Second image with caption (figure)
- `image_003.jpg` - Third image (format preserved)

### Using in Sphinx

1. Extract the ZIP to your Sphinx project's source directory
2. Rename `document.rst` as needed
3. Add to your `toctree` in `index.rst`
4. Build your documentation: `make html`

---

## Troubleshooting

### Preview Not Updating

- Click the **Refresh** button to manually update
- Check if the document has unsaved changes
- Try closing and reopening the preview panel

### Images Missing in Export

- Ensure images are inserted as inline pictures (not floating)
- Very large images may take time to process
- Check browser console for any error messages

### Custom Styles Not Converting

- Style name must start with `rst_` (case-sensitive)
- Ensure the style is applied to the entire paragraph
- Check for typos in the style name

### Export Download Fails

- Check your browser's download settings
- Ensure pop-ups are not blocked
- Try a different browser

### Add-in Not Loading

- Clear browser cache and reload
- Check that JavaScript is enabled
- Try sideloading the manifest again
- Ensure you're using a supported browser (Edge, Chrome, Firefox)

---

## Feedback and Support

- **Report Issues**: [GitHub Issues](https://github.com/your-repo/rst-word-addin/issues)
- **Feature Requests**: [GitHub Discussions](https://github.com/your-repo/rst-word-addin/discussions)

---

## Next Steps

- [FORMATTING.md](FORMATTING.md) - Complete formatting reference
- [CUSTOM_STYLES.md](CUSTOM_STYLES.md) - Creating custom directive styles
- [EXAMPLES.md](EXAMPLES.md) - Example documents and conversions
