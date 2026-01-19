# RST Formatting Reference

Complete reference for how Word formatting converts to reStructuredText.

## Table of Contents

- [Text Formatting](#text-formatting)
- [Headings](#headings)
- [Paragraphs](#paragraphs)
- [Lists](#lists)
- [Links](#links)
- [Images](#images)
- [Figures](#figures)
- [Tables](#tables)
- [Table of Contents](#table-of-contents)
- [Special Characters](#special-characters)
- [Unsupported Formatting](#unsupported-formatting)

---

## Text Formatting

### Bold

| Word | RST |
|------|-----|
| **Bold text** | `**Bold text**` |

Select text and press `Ctrl+B` or use the Bold button.

### Italic

| Word | RST |
|------|-----|
| *Italic text* | `*Italic text*` |

Select text and press `Ctrl+I` or use the Italic button.

### Bold and Italic

| Word | RST |
|------|-----|
| ***Bold italic*** | `***Bold italic***` |

Apply both bold and italic formatting to the same text.

### Inline Code

| Word | RST |
|------|-----|
| `monospace text` | ``` ``monospace text`` ``` |

Apply a monospace font (Consolas, Courier New, etc.) to create inline code.

**Tip:** Create a character style called "Code" with Consolas font for easy application.

### Subscript and Superscript

| Word | RST |
|------|-----|
| H₂O | `:sub:`2`` → H\ :sub:`2`\ O |
| E=mc² | `:sup:`2`` → mc\ :sup:`2` |

Note: RST subscript/superscript requires escaping adjacent text with backslashes.

---

## Headings

Word heading styles convert to RST section titles with underline (and optional overline) characters.

### Heading Hierarchy

| Word Style | RST Format | Example |
|------------|------------|---------|
| **Heading 1** | Overline + underline with `=` | Document title |
| **Heading 2** | Underline with `=` | Major sections |
| **Heading 3** | Underline with `-` | Subsections |
| **Heading 4** | Underline with `~` | Sub-subsections |
| **Heading 5** | Underline with `^` | Paragraphs |
| **Heading 6** | Underline with `"` | Minor headings |

### Example Output

**Word:**
- Heading 1: "User Guide"
- Heading 2: "Getting Started"
- Heading 3: "Installation"

**RST:**
```rst
==========
User Guide
==========

Getting Started
===============

Installation
------------
```

### Heading Rules

1. Underline must be **at least as long** as the heading text
2. Use consistent characters throughout the document
3. RST determines hierarchy by the **order of appearance**, not the character used
4. The add-in maintains consistent hierarchy based on Word heading levels

---

## Paragraphs

### Normal Paragraphs

Regular paragraphs convert directly. Blank lines separate paragraphs in RST.

**Word:**
```
This is the first paragraph with some text.

This is the second paragraph.
```

**RST:**
```rst
This is the first paragraph with some text.

This is the second paragraph.
```

### Line Breaks

Soft line breaks (Shift+Enter) within a paragraph are preserved.

**Word:**
```
Line one
Line two (same paragraph)
```

**RST:**
```rst
| Line one
| Line two (same paragraph)
```

### Block Quotes

Indented paragraphs convert to RST block quotes.

**Word:** (Paragraph with left indent)
```
    This is an indented block quote.
    It can span multiple lines.
```

**RST:**
```rst
   This is an indented block quote.
   It can span multiple lines.
```

---

## Lists

### Bulleted Lists

**Word:**
- First item
- Second item
- Third item

**RST:**
```rst
- First item
- Second item
- Third item
```

### Numbered Lists

**Word:**
1. First item
2. Second item
3. Third item

**RST:**
```rst
1. First item
2. Second item
3. Third item
```

### Nested Lists

**Word:**
- Item one
  - Nested item A
  - Nested item B
- Item two

**RST:**
```rst
- Item one

  - Nested item A
  - Nested item B

- Item two
```

### Mixed Lists

**Word:**
1. First numbered
   - Bullet under numbered
   - Another bullet
2. Second numbered

**RST:**
```rst
1. First numbered

   - Bullet under numbered
   - Another bullet

2. Second numbered
```

### Definition Lists

Use bold text followed by a colon at the start of a paragraph, then the definition on the next line (indented).

**Word:**
```
**Term:**
    Definition of the term goes here.

**Another term:**
    Another definition.
```

**RST:**
```rst
Term
   Definition of the term goes here.

Another term
   Another definition.
```

---

## Links

### Hyperlinks

**Word:** Text with a hyperlink applied

**RST:**
```rst
`Link text <https://example.com>`__
```

### Email Links

**Word:** email@example.com (hyperlinked)

**RST:**
```rst
`email@example.com <mailto:email@example.com>`__
```

### Internal Document Links

Bookmarks in Word can convert to RST internal references:

**RST:**
```rst
See the :ref:`installation` section for details.

.. _installation:

Installation
------------
```

---

## Images

### Basic Images

Images inserted in Word convert to the `.. image::` directive.

**RST:**
```rst
.. image:: images/screenshot.png
```

### Image with Attributes

The add-in extracts available attributes from Word:

**RST:**
```rst
.. image:: images/diagram.png
   :alt: Architecture diagram
   :width: 500px
   :height: 300px
   :align: center
```

### Supported Image Attributes

| Attribute | Source | Description |
|-----------|--------|-------------|
| `:alt:` | Word alt text | Alternative text for accessibility |
| `:width:` | Word image width | Display width |
| `:height:` | Word image height | Display height |
| `:scale:` | Calculated | Scale percentage |
| `:align:` | Word alignment | `left`, `center`, or `right` |
| `:target:` | Word hyperlink | Makes image clickable |

### Setting Alt Text in Word

1. Right-click the image
2. Select **Edit Alt Text** (or **View Alt Text**)
3. Enter descriptive text
4. This becomes the `:alt:` attribute

### Image Alignment

| Word Alignment | RST `:align:` |
|----------------|---------------|
| Left | `left` |
| Center | `center` |
| Right | `right` |
| Inline with text | (no align attribute) |

---

## Figures

Images with captions use the `.. figure::` directive instead of `.. image::`.

### Adding Captions in Word

1. Select the image
2. Go to **References** tab
3. Click **Insert Caption**
4. Choose "Figure" label and add description
5. Click **OK**

### Figure Output

**Word:** Image with caption "Figure 1: System Architecture"

**RST:**
```rst
.. figure:: images/architecture.png
   :alt: System Architecture
   :width: 600px
   :align: center

   Figure 1: System Architecture
```

### Figure-Specific Attributes

| Attribute | Description |
|-----------|-------------|
| `:figwidth:` | Width of figure container |
| `:figclass:` | CSS class for figure element |

### Figure with Legend

If additional text follows the caption (in the same text box or immediately after), it becomes a legend:

**RST:**
```rst
.. figure:: images/workflow.png
   :align: center

   Figure 2: Data Processing Workflow

   This diagram shows the complete data flow from input
   to output, including all transformation steps.
```

---

## Tables

### Basic Tables

Word tables convert to RST grid table format.

**Word:**

| Name | Age | City |
|------|-----|------|
| Alice | 30 | NYC |
| Bob | 25 | LA |

**RST:**
```rst
+-------+-----+------+
| Name  | Age | City |
+=======+=====+======+
| Alice | 30  | NYC  |
+-------+-----+------+
| Bob   | 25  | LA   |
+-------+-----+------+
```

### Tables with Captions

Add a caption using **References** → **Insert Caption** → "Table" label.

**RST:**
```rst
.. table:: Table 1: User Information
   :align: center

   +-------+-----+------+
   | Name  | Age | City |
   +=======+=====+======+
   | Alice | 30  | NYC  |
   +-------+-----+------+
```

### Table Attributes

| Attribute | Description |
|-----------|-------------|
| `:align:` | Table alignment: `left`, `center`, `right` |
| `:width:` | Total table width |
| `:widths:` | Column width ratios |

### Header Rows

The first row of a Word table (if formatted differently) becomes the header row, indicated by `=` instead of `-` in the separator.

### Merged Cells

RST grid tables support cell spanning:

**RST:**
```rst
+-------+-------+-------+
| Header spanning 3 cols|
+-------+-------+-------+
| A     | B     | C     |
+-------+-------+-------+
```

Note: Complex merged cells may require manual adjustment.

---

## Table of Contents

### Word TOC

If your Word document contains a Table of Contents field, it converts to the RST `.. contents::` directive.

**RST:**
```rst
.. contents:: Table of Contents
   :depth: 3
```

### TOC Attributes

| Attribute | Description |
|-----------|-------------|
| `:depth:` | Maximum heading levels to include |
| `:local:` | Only show subsections of current section |
| `:backlinks:` | `entry`, `top`, or `none` |

### Creating TOC in Word

1. Place cursor where you want the TOC
2. Go to **References** tab
3. Click **Table of Contents**
4. Choose a style or **Custom Table of Contents**
5. Set the number of levels to show

---

## Special Characters

### Character Escaping

RST uses certain characters for markup. If these appear in your Word text, they're automatically escaped:

| Character | Escaped As | Meaning in RST |
|-----------|------------|----------------|
| `*` | `\*` | Emphasis markers |
| `` ` `` | `` \` `` | Inline literal |
| `_` | `\_` | Hyperlink/reference |
| `\` | `\\` | Escape character |
| `|` | `\|` | Substitution |

### Unicode Characters

Unicode characters in Word are preserved in RST output. Ensure your RST file uses UTF-8 encoding.

### Special Symbols

| Word | RST |
|------|-----|
| — (em dash) | `—` or `---` |
| – (en dash) | `–` or `--` |
| … (ellipsis) | `…` or `...` |
| © | `©` |
| ® | `®` |
| ™ | `™` |

---

## Unsupported Formatting

The following Word features have limited or no RST equivalent:

### Not Supported

| Word Feature | Notes |
|--------------|-------|
| Underline | No standard RST equivalent |
| Strikethrough | No standard RST equivalent |
| Text color | RST is plain text |
| Highlight color | RST is plain text |
| Font changes | RST doesn't specify fonts |
| Text boxes | Content extracted, positioning lost |
| SmartArt | Converted to text if possible |
| Charts | Not supported - use images |
| Equations | Limited support - consider LaTeX |

### Partial Support

| Word Feature | RST Handling |
|--------------|--------------|
| Footnotes | Converted to RST footnotes `[#]_` |
| Endnotes | Converted to RST footnotes |
| Comments | Ignored (not included in output) |
| Track changes | Final text only (changes not shown) |
| Headers/Footers | Ignored |
| Page breaks | Ignored |
| Columns | Content linearized |

### Recommendations

1. **Text color/highlight**: Use admonition directives (`rst_note`, `rst_warning`) instead
2. **Charts**: Export as image, then insert the image
3. **Equations**: Use `rst_math` style with LaTeX syntax
4. **Complex layouts**: Simplify before converting

---

## Best Practices

### For Best Results

1. **Use built-in styles** - Heading 1-6, Normal, List Bullet, List Number
2. **Add alt text to images** - Improves accessibility and RST output
3. **Use Insert Caption** - For proper figure/table numbering
4. **Keep tables simple** - Avoid heavily merged cells
5. **Use hyperlinks properly** - Select text, then add link

### Document Preparation Checklist

Before converting your document:

- [ ] Apply heading styles consistently
- [ ] Add alt text to all images
- [ ] Add captions to figures and tables
- [ ] Check that lists use proper list styles
- [ ] Remove or replace unsupported formatting
- [ ] Preview the RST output to catch issues

---

## See Also

- [README.md](README.md) - Main user guide
- [CUSTOM_STYLES.md](CUSTOM_STYLES.md) - Creating custom directive styles
- [reStructuredText Primer](https://www.sphinx-doc.org/en/master/usage/restructuredtext/basics.html)
- [RST Specification](https://docutils.sourceforge.io/docs/ref/rst/restructuredtext.html)
