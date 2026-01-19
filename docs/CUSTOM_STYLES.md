# Custom RST Directive Styles Guide

This guide explains how to create and use custom Word styles that convert to RST directives.

## Table of Contents

- [Overview](#overview)
- [Creating Custom Styles in Word](#creating-custom-styles-in-word)
- [Style Naming Convention](#style-naming-convention)
- [Content Format](#content-format)
- [Directive Reference](#directive-reference)
- [Examples](#examples)
- [Tips and Best Practices](#tips-and-best-practices)

---

## Overview

The RST Word Add-in converts Word paragraph styles that begin with `rst_` into RST directives. This allows you to use any RST directive - standard or custom - directly in your Word documents.

**How it works:**

| Word Style Name | RST Directive |
|-----------------|---------------|
| `rst_note` | `.. note::` |
| `rst_warning` | `.. warning::` |
| `rst_code-block` | `.. code-block::` |
| `rst_need` | `.. need::` |
| `rst_my-custom` | `.. my-custom::` |

---

## Creating Custom Styles in Word

### Step-by-Step Instructions

#### Word Online

1. Select some text in your document
2. Go to **Home** tab
3. In the **Styles** group, click the **More** button (↓)
4. Click **Create a Style**
5. Enter the style name (e.g., `rst_note`)
6. Click **Modify** to customize appearance (optional)
7. Click **OK**

#### Word Desktop (Windows)

1. Go to **Home** tab
2. Click the small arrow (↘) in the **Styles** group to open the Styles pane
3. Click the **New Style** button (A with sparkle) at bottom
4. Configure:
   - **Name**: `rst_note` (or your directive name)
   - **Style type**: Paragraph
   - **Style based on**: Normal
   - **Style for following paragraph**: Normal
5. Click **Format** to customize font, paragraph, borders, etc.
6. Click **OK**

#### Word Desktop (Mac)

1. Go to **Home** tab
2. Click the **Styles Pane** button
3. Click **New Style** at bottom of pane
4. Configure:
   - **Name**: `rst_note`
   - **Style type**: Paragraph
5. Customize formatting as desired
6. Click **OK**

### Recommended Style Formatting

To make directive content visually distinct in Word, consider:

| Property | Recommendation |
|----------|----------------|
| **Background** | Light gray (#f0f0f0) or light color |
| **Left border** | 3pt solid colored border |
| **Left indent** | 0.25" - 0.5" |
| **Font** | Monospace for code (Consolas, Courier New) |
| **Font size** | Same or slightly smaller than body text |

**Example - Note Style:**
- Background: Light yellow (#ffffd0)
- Left border: 3pt solid gold
- Left indent: 0.25"
- Paragraph spacing: 6pt before, 6pt after

**Example - Code Block Style:**
- Background: Light gray (#f5f5f5)
- Font: Consolas 10pt
- Left border: 3pt solid blue
- No first line indent

---

## Style Naming Convention

### Basic Rules

1. Style name **must** start with `rst_` (lowercase)
2. After `rst_`, use the exact directive name
3. Hyphens in directive names stay as hyphens: `rst_code-block`
4. Case matters: `rst_Note` will NOT work, use `rst_note`

### Examples

| Directive | Style Name |
|-----------|------------|
| `.. note::` | `rst_note` |
| `.. warning::` | `rst_warning` |
| `.. code-block::` | `rst_code-block` |
| `.. pull-quote::` | `rst_pull-quote` |
| `.. versionadded::` | `rst_versionadded` |
| `.. todo::` | `rst_todo` |
| `.. need::` | `rst_need` |

---

## Content Format

### Simple Content (No Arguments)

Just type your content with the `rst_*` style applied:

**Word (style: `rst_note`):**
```
This is an important note that readers should pay attention to.
```

**RST Output:**
```rst
.. note::

   This is an important note that readers should pay attention to.
```

### Content with Directive Argument

If the directive needs an argument, put it on the first line in square brackets:

**Word (style: `rst_code-block`):**
```
[python]
def greet(name):
    print(f"Hello, {name}!")
```

**RST Output:**
```rst
.. code-block:: python

   def greet(name):
       print(f"Hello, {name}!")
```

### Content with Options

Add options using `:option: value` format, one per line:

**Word (style: `rst_image`):**
```
[images/logo.png]
:alt: Company Logo
:width: 200px
:align: center
```

**RST Output:**
```rst
.. image:: images/logo.png
   :alt: Company Logo
   :width: 200px
   :align: center
```

### Content with Arguments, Options, AND Body

Combine all three - argument first, then options, then body content:

**Word (style: `rst_admonition`):**
```
[Custom Title Here]
:class: my-custom-class
This is the body content of the admonition.

It can span multiple paragraphs.
```

**RST Output:**
```rst
.. admonition:: Custom Title Here
   :class: my-custom-class

   This is the body content of the admonition.

   It can span multiple paragraphs.
```

### Format Summary

```
[argument]              ← Line 1: Optional, in square brackets
:option1: value1        ← Lines 2+: Optional, start with colon
:option2: value2
                        ← Blank line (optional in Word)
Body content here       ← Remaining lines become directive body
More body content
```

---

## Directive Reference

### Admonitions

Standard RST admonition directives:

| Style Name | Purpose | Icon Color |
|------------|---------|------------|
| `rst_note` | General notes | Blue |
| `rst_warning` | Warnings | Orange |
| `rst_danger` | Critical warnings | Red |
| `rst_error` | Error messages | Red |
| `rst_hint` | Helpful hints | Green |
| `rst_tip` | Tips and tricks | Green |
| `rst_important` | Important info | Orange |
| `rst_caution` | Caution notices | Yellow |
| `rst_attention` | Attention notices | Yellow |
| `rst_admonition` | Custom title admonition | — |

**Example - Warning:**

Word (style: `rst_warning`):
```
Do not delete system files! This may cause data loss.
```

RST Output:
```rst
.. warning::

   Do not delete system files! This may cause data loss.
```

### Code Blocks

| Style Name | Purpose |
|------------|---------|
| `rst_code-block` | Syntax-highlighted code |
| `rst_code` | Simple code block |
| `rst_literalinclude` | Include code from file |

**Common Options:**

| Option | Description |
|--------|-------------|
| `:linenos:` | Show line numbers |
| `:emphasize-lines:` | Highlight specific lines (e.g., `1,3,5-7`) |
| `:caption:` | Add a caption above the code |
| `:name:` | Reference label for cross-references |

**Example - Python with Line Numbers:**

Word (style: `rst_code-block`):
```
[python]
:linenos:
:caption: Hello World Example
def main():
    print("Hello, World!")

if __name__ == "__main__":
    main()
```

RST Output:
```rst
.. code-block:: python
   :linenos:
   :caption: Hello World Example

   def main():
       print("Hello, World!")

   if __name__ == "__main__":
       main()
```

**Supported Languages:**

`python`, `javascript`, `typescript`, `java`, `c`, `cpp`, `csharp`, `go`, `rust`, `ruby`, `php`, `sql`, `bash`, `powershell`, `json`, `yaml`, `xml`, `html`, `css`, `markdown`, `rst`, and many more.

### Sphinx-Needs Directives

For requirements management with [Sphinx-Needs](https://sphinx-needs.readthedocs.io/):

| Style Name | Purpose |
|------------|---------|
| `rst_need` | Generic need |
| `rst_req` | Requirement |
| `rst_spec` | Specification |
| `rst_impl` | Implementation |
| `rst_test` | Test case |

**Common Options:**

| Option | Description |
|--------|-------------|
| `:id:` | Unique identifier (required) |
| `:title:` | Display title |
| `:status:` | Status (open, in progress, closed) |
| `:tags:` | Comma-separated tags |
| `:links:` | Links to other needs |

**Example - Requirement:**

Word (style: `rst_req`):
```
:id: REQ-001
:title: User Authentication
:status: open
:tags: security, login
:links: SPEC-001, SPEC-002
The system shall authenticate users using OAuth 2.0 or SAML 2.0 protocols.

Accepted authentication providers:
- Google
- Microsoft Azure AD
- Okta
```

RST Output:
```rst
.. req::
   :id: REQ-001
   :title: User Authentication
   :status: open
   :tags: security, login
   :links: SPEC-001, SPEC-002

   The system shall authenticate users using OAuth 2.0 or SAML 2.0 protocols.

   Accepted authentication providers:
   - Google
   - Microsoft Azure AD
   - Okta
```

### Structure Directives

| Style Name | Purpose |
|------------|---------|
| `rst_topic` | Topic block with title |
| `rst_sidebar` | Sidebar content |
| `rst_rubric` | Informal heading |
| `rst_epigraph` | Quotation block |
| `rst_pull-quote` | Pull quote |
| `rst_highlights` | Summary highlights |
| `rst_compound` | Compound paragraph |

### Sphinx Directives

| Style Name | Purpose |
|------------|---------|
| `rst_toctree` | Table of contents tree |
| `rst_only` | Conditional content |
| `rst_index` | Index entries |
| `rst_glossary` | Glossary terms |
| `rst_deprecated` | Deprecation notice |
| `rst_versionadded` | Version added notice |
| `rst_versionchanged` | Version changed notice |
| `rst_seealso` | See also block |
| `rst_todo` | TODO items (with sphinx.ext.todo) |

**Example - Version Added:**

Word (style: `rst_versionadded`):
```
[2.0]
The export feature now supports ZIP compression.
```

RST Output:
```rst
.. versionadded:: 2.0

   The export feature now supports ZIP compression.
```

---

## Examples

### Complete Document Example

Here's how a Word document might use custom styles:

---

**[Heading 1 style]** API Reference

**[Normal style]** This section documents the REST API endpoints.

**[rst_warning style]**
```
The API is currently in beta. Breaking changes may occur.
```

**[Heading 2 style]** Authentication

**[Normal style]** All API requests require authentication.

**[rst_code-block style]**
```
[bash]
curl -H "Authorization: Bearer TOKEN" https://api.example.com/v1/users
```

**[rst_note style]**
```
Tokens expire after 24 hours. Use the refresh endpoint to obtain a new token.
```

**[Heading 2 style]** Requirements

**[rst_req style]**
```
:id: API-001
:title: Rate Limiting
:status: implemented
The API shall limit requests to 100 per minute per user.
```

---

**RST Output:**

```rst
=============
API Reference
=============

This section documents the REST API endpoints.

.. warning::

   The API is currently in beta. Breaking changes may occur.

Authentication
==============

All API requests require authentication.

.. code-block:: bash

   curl -H "Authorization: Bearer TOKEN" https://api.example.com/v1/users

.. note::

   Tokens expire after 24 hours. Use the refresh endpoint to obtain a new token.

Requirements
============

.. req::
   :id: API-001
   :title: Rate Limiting
   :status: implemented

   The API shall limit requests to 100 per minute per user.
```

---

## Tips and Best Practices

### Style Organization

1. **Create a style set** for your project with all needed `rst_*` styles
2. **Save as template** (.dotx) to reuse across documents
3. **Use consistent formatting** so directives are visually distinct

### Content Tips

1. **Keep arguments simple** - no special characters in `[argument]`
2. **One option per line** - makes it easier to read and edit
3. **Blank lines in body** - Use Word paragraph breaks, they convert correctly

### Debugging

If your directive isn't converting correctly:

1. **Check style name** - Must start with `rst_` (lowercase)
2. **Check style application** - Apply to entire paragraph
3. **Check argument format** - Must be `[argument]` on first line
4. **Check option format** - Must be `:option: value` with colon at start and end of option name

### Style Template

Here's a recommended set of styles to create for technical documentation:

```
Essential:
- rst_note
- rst_warning
- rst_tip
- rst_code-block

For APIs:
- rst_deprecated
- rst_versionadded
- rst_versionchanged

For Requirements:
- rst_req
- rst_spec
- rst_test

For Structure:
- rst_topic
- rst_seealso
- rst_todo
```

---

## See Also

- [README.md](README.md) - Main user guide
- [FORMATTING.md](FORMATTING.md) - Complete formatting reference
- [Sphinx Directives](https://www.sphinx-doc.org/en/master/usage/restructuredtext/directives.html)
- [Sphinx-Needs Documentation](https://sphinx-needs.readthedocs.io/)
