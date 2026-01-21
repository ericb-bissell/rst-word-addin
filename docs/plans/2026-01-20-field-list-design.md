# Field List Support Design

**Date:** 2026-01-20
**Feature:** RST Field List conversion from Word documents

## Overview

Add support for converting tab-separated field definitions in Word documents to RST field lists.

## Input Format

Word document content using the pattern:

```
FieldName::\tValue
FieldName::
```

Where:
- Field name is a single word (no spaces), followed by double colon (`::`)
- Tab character (`\t`) separates name from value (optional if value is empty)
- Line break (LF or CR/LF) terminates each field
- Consecutive matching lines form a single field list

**Example Word input:**
```
Author::	John Doe
Version::	1.0
Status::
Priority::	High
```

## Output Format

RST field list syntax:

```rst
:Author: John Doe
:Version: 1.0
:Status:
:Priority: High
```

Field lists are separated from surrounding content by blank lines.

## Detection Pattern

Regex: `/^(\w+)::(?:\t(.*))?$/`

- `(\w+)` - Captures field name (word characters only, no spaces)
- `::` - Literal double colon marker
- `(?:\t(.*))?` - Optional: tab followed by value (captured)

## Implementation

### 1. Types (`src/converter/types.ts`)

```typescript
export interface FieldListItem {
  name: string;
  value: string;
}

export interface FieldListElement extends DocumentElement {
  type: 'field-list';
  fields: FieldListItem[];
}
```

Add `'field-list'` to `DocumentElementType` union.

### 2. Parser (`src/converter/html-parser.ts`)

Add detection function:

```typescript
function parseFieldListLine(content: string): FieldListItem | null {
  const match = content.match(/^(\w+)::(?:\t(.*))?$/);
  if (match) {
    return { name: match[1], value: match[2] || '' };
  }
  return null;
}
```

Modify paragraph processing:
- Check each paragraph for field list pattern
- Collect consecutive field lines into a `FieldListElement`
- Non-matching lines break the field list sequence

### 3. Formatter (`src/converter/rst-formatter.ts`)

Add formatting function:

```typescript
function formatFieldList(element: FieldListElement): string {
  return element.fields
    .map(field => `:${field.name}: ${field.value}`.trimEnd())
    .join('\n');
}
```

Add case in `formatElement()` switch statement.

## Edge Cases

| Input | Output | Notes |
|-------|--------|-------|
| `Status::` | `:Status:` | Empty value, no tab needed |
| `Status::\t` | `:Status:` | Empty value with tab |
| `Name::\tValue` | `:Name: Value` | Standard case |
| Single field line | Still creates field list | No minimum count |
| `Invalid Name::\tX` | Regular paragraph | Space in name = not a field |
| `Name:\tValue` | Regular paragraph | Single colon = not a field |

## Files Modified

1. `src/converter/types.ts` - Add types
2. `src/converter/html-parser.ts` - Add detection and parsing
3. `src/converter/rst-formatter.ts` - Add formatting
4. `CLAUDE.md` - Update version and feature status

## Testing

Test cases:
- Single field
- Multiple consecutive fields
- Field with empty value
- Field list interrupted by regular paragraph
- Text that looks similar but isn't a field (single colon, spaces in name)
