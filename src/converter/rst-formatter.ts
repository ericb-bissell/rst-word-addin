/**
 * RST Word Add-in - RST Formatter
 * Converts document elements into properly formatted RST text
 *
 * Handles all RST formatting rules including:
 * - Heading underlines with proper hierarchy
 * - Text formatting (bold, italic, code)
 * - Lists (ordered, unordered, nested)
 * - Block quotes
 * - References and links
 */

import {
  AnyDocumentElement,
  HeadingElement,
  ParagraphElement,
  ListElement,
  ImageElement,
  FigureElement,
  TableElement,
  TocElement,
  DirectiveElement,
} from './types';

import {
  generateImageDirective,
  generateFigureDirective,
  generateTableDirective,
  generateContentsDirective,
  generateCustomDirective,
} from './directives';

/**
 * RST heading characters in order of precedence
 * Different characters create different heading levels
 */
const HEADING_CHARS = ['=', '-', '~', '^', '"', "'"];

/**
 * Formatter options
 */
export interface FormatterOptions {
  /** Number of columns for line wrapping (0 = no wrap) */
  lineWidth: number;
  /** Use overline for title (level 1) */
  titleOverline: boolean;
  /** Indent size for directive content */
  indentSize: number;
  /** Image directory path prefix */
  imageDir: string;
}

/**
 * Default formatter options
 */
const DEFAULT_OPTIONS: FormatterOptions = {
  lineWidth: 0, // No wrapping by default
  titleOverline: true,
  indentSize: 3,
  imageDir: 'images/',
};

/**
 * Format a single document element to RST
 *
 * @param element - Document element to format
 * @param options - Formatter options
 * @returns RST formatted string
 */
export function formatElement(
  element: AnyDocumentElement,
  options: Partial<FormatterOptions> = {}
): string {
  const opts = { ...DEFAULT_OPTIONS, ...options };

  switch (element.type) {
    case 'heading':
      return formatHeading(element as HeadingElement, opts);
    case 'paragraph':
      return formatParagraph(element as ParagraphElement, opts);
    case 'list':
      return formatList(element as ListElement, opts);
    case 'image':
      return formatImage(element as ImageElement, opts);
    case 'figure':
      return formatFigure(element as FigureElement, opts);
    case 'table':
      return formatTable(element as TableElement, opts);
    case 'toc':
      return formatToc(element as TocElement, opts);
    case 'directive':
      return formatDirective(element as DirectiveElement, opts);
    default:
      return '';
  }
}

/**
 * Format multiple elements to RST document
 *
 * @param elements - Document elements to format
 * @param options - Formatter options
 * @returns Complete RST document string
 */
export function formatDocument(
  elements: AnyDocumentElement[],
  options: Partial<FormatterOptions> = {}
): string {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const parts: string[] = [];

  for (let i = 0; i < elements.length; i++) {
    const element = elements[i];
    const formatted = formatElement(element, opts);

    if (formatted) {
      parts.push(formatted);
    }
  }

  // Join with appropriate spacing
  return joinWithSpacing(parts);
}

/**
 * Join RST parts with appropriate blank line spacing
 */
function joinWithSpacing(parts: string[]): string {
  if (parts.length === 0) {
    return '';
  }

  const result: string[] = [];

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i].trim();
    if (!part) continue;

    if (i > 0) {
      // Add blank line between most elements
      result.push('');
    }

    result.push(part);
  }

  return result.join('\n');
}

/**
 * Format heading element
 *
 * RST heading levels are indicated by underlines (and optional overlines):
 * - Level 1 (title): = with overline
 * - Level 2: = underline only
 * - Level 3: - underline only
 * - Level 4: ~ underline only
 * - etc.
 */
function formatHeading(element: HeadingElement, options: FormatterOptions): string {
  const { level, text } = element;
  const charIndex = Math.min(level - 1, HEADING_CHARS.length - 1);
  const underlineChar = HEADING_CHARS[charIndex];
  const underline = underlineChar.repeat(getTextWidth(text));

  const lines: string[] = [];

  // Add overline for title (level 1) if enabled
  if (level === 1 && options.titleOverline) {
    lines.push(underline);
  }

  lines.push(text);
  lines.push(underline);

  return lines.join('\n');
}

/**
 * Format paragraph element
 */
function formatParagraph(element: ParagraphElement, options: FormatterOptions): string {
  let content = element.content;

  // Apply line wrapping if configured
  if (options.lineWidth > 0) {
    content = wrapText(content, options.lineWidth);
  }

  // Handle blockquote
  if (element.isBlockQuote) {
    return indentText(content, options.indentSize);
  }

  return content;
}

/**
 * Format list element
 */
function formatList(element: ListElement, options: FormatterOptions, depth: number = 0): string {
  const { listType, items } = element;
  const lines: string[] = [];
  const indent = '  '.repeat(depth); // RST uses 2-space indent for nested lists

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    // Use #. for auto-numbered ordered lists in RST
    const marker = listType === 'ordered' ? '#.' : '-';

    // First line with marker
    const firstLine = `${indent}${marker} ${item.content}`;

    // Apply wrapping if configured
    if (options.lineWidth > 0) {
      const wrappedContent = wrapText(item.content, options.lineWidth - indent.length - marker.length - 1);
      const contentLines = wrappedContent.split('\n');
      lines.push(`${indent}${marker} ${contentLines[0]}`);

      // Continuation lines need extra indent (align with content after marker)
      const continuationIndent = indent + ' '.repeat(marker.length + 1);
      for (let j = 1; j < contentLines.length; j++) {
        lines.push(`${continuationIndent}${contentLines[j]}`);
      }
    } else {
      lines.push(firstLine);
    }

    // Handle nested list - no blank line needed, just indent
    if (item.nestedList) {
      lines.push(formatList(item.nestedList, options, depth + 1));
    }
  }

  return lines.join('\n');
}

/**
 * Format image element using image directive
 */
function formatImage(element: ImageElement, _options: FormatterOptions): string {
  const imageOptions = { ...element.options };

  // Adjust image path if needed
  if (element.imageData && !imageOptions.uri.startsWith('http')) {
    imageOptions.uri = element.imageData.filename;
  }

  return generateImageDirective(imageOptions);
}

/**
 * Format figure element using figure directive
 */
function formatFigure(element: FigureElement, _options: FormatterOptions): string {
  const figureOptions = { ...element.options };

  // Adjust image path if needed
  if (element.imageData && !figureOptions.uri.startsWith('http')) {
    figureOptions.uri = element.imageData.filename;
  }

  return generateFigureDirective(figureOptions);
}

/**
 * Format table element using table directive
 */
function formatTable(element: TableElement, _options: FormatterOptions): string {
  return generateTableDirective(element.data);
}

/**
 * Format TOC element using contents directive
 */
function formatToc(element: TocElement, _options: FormatterOptions): string {
  return generateContentsDirective(element.options);
}

/**
 * Format custom directive element
 */
function formatDirective(element: DirectiveElement, _options: FormatterOptions): string {
  return generateCustomDirective(element.directive);
}

/**
 * Get display width of text (handles some Unicode)
 */
function getTextWidth(text: string): number {
  // Simple implementation - could be enhanced for full Unicode support
  return text.length;
}

/**
 * Wrap text to specified width
 */
function wrapText(text: string, width: number): string {
  if (width <= 0 || text.length <= width) {
    return text;
  }

  const words = text.split(/\s+/);
  const lines: string[] = [];
  let currentLine = '';

  for (const word of words) {
    if (currentLine.length === 0) {
      currentLine = word;
    } else if (currentLine.length + 1 + word.length <= width) {
      currentLine += ' ' + word;
    } else {
      lines.push(currentLine);
      currentLine = word;
    }
  }

  if (currentLine) {
    lines.push(currentLine);
  }

  return lines.join('\n');
}

/**
 * Indent text by specified amount
 */
function indentText(text: string, spaces: number): string {
  const indent = ' '.repeat(spaces);
  return text
    .split('\n')
    .map((line) => (line.trim() ? indent + line : ''))
    .join('\n');
}

/**
 * Escape special RST characters in text
 */
export function escapeRstText(text: string): string {
  // Characters that need escaping in certain contexts
  const specialChars = ['*', '`', '|', '_'];

  let result = text;

  // Escape backslashes first
  result = result.replace(/\\/g, '\\\\');

  // Escape special chars at word boundaries (simplified approach)
  for (const char of specialChars) {
    const regex = new RegExp(`(^|\\s)\\${char}|\\${char}($|\\s)`, 'g');
    result = result.replace(regex, (match) => {
      return match.replace(char, '\\' + char);
    });
  }

  return result;
}

/**
 * Create an RST reference/label
 */
export function createLabel(name: string): string {
  // Normalize to valid label format
  const normalized = name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');

  return `.. _${normalized}:`;
}

/**
 * Create an RST inline reference
 */
export function createRef(label: string, text?: string): string {
  if (text) {
    return `:ref:\`${text} <${label}>\``;
  }
  return `:ref:\`${label}\``;
}

/**
 * Create an RST external link
 */
export function createLink(url: string, text: string): string {
  return `\`${text} <${url}>\`_`;
}

/**
 * Create RST substitution definition
 */
export function createSubstitution(name: string, replacement: string): string {
  return `.. |${name}| replace:: ${replacement}`;
}

/**
 * Create RST comment
 */
export function createComment(text: string): string {
  const lines = text.split('\n');
  if (lines.length === 1) {
    return `.. ${text}`;
  }

  return ['..', ...lines.map((line) => '   ' + line)].join('\n');
}

/**
 * Format inline code
 */
export function formatInlineCode(code: string): string {
  return `\`\`${code}\`\``;
}

/**
 * Format bold text
 */
export function formatBold(text: string): string {
  return `**${text}**`;
}

/**
 * Format italic text
 */
export function formatItalic(text: string): string {
  return `*${text}*`;
}

/**
 * Format literal/code block
 */
export function formatCodeBlock(code: string, language?: string): string {
  const lines: string[] = [];

  if (language) {
    lines.push(`.. code-block:: ${language}`);
  } else {
    lines.push('::');
  }

  lines.push('');

  // Indent code
  const codeLines = code.split('\n');
  for (const line of codeLines) {
    lines.push('   ' + line);
  }

  return lines.join('\n');
}

/**
 * Format a field list entry
 */
export function formatField(name: string, value: string): string {
  return `:${name}: ${value}`;
}

/**
 * Format definition list item
 */
export function formatDefinition(term: string, definition: string): string {
  const defLines = definition.split('\n');
  const indented = defLines.map((line) => '   ' + line).join('\n');
  return `${term}\n${indented}`;
}
