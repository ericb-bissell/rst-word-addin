/**
 * RST Word Add-in - Shared Types
 * Type definitions for the converter modules
 */

/**
 * RST image directive options
 * @see https://docutils.sourceforge.io/docs/ref/rst/directives.html#image
 */
export interface ImageOptions {
  /** Path to the image file */
  uri: string;
  /** Alternative text for accessibility */
  alt?: string;
  /** Image height (e.g., "200px", "5cm") */
  height?: string;
  /** Image width (e.g., "400px", "80%") */
  width?: string;
  /** Uniform scaling factor as percentage (e.g., "75") */
  scale?: string;
  /** Alignment: inline (top, middle, bottom) or block (left, center, right) */
  align?: 'top' | 'middle' | 'bottom' | 'left' | 'center' | 'right';
  /** Makes image clickable, links to this URI */
  target?: string;
  /** CSS class names (space-separated) */
  class?: string;
  /** Reference name for cross-references */
  name?: string;
  /** Loading behavior: embed, link, or lazy */
  loading?: 'embed' | 'link' | 'lazy';
}

/**
 * RST figure directive options (extends image options)
 * @see https://docutils.sourceforge.io/docs/ref/rst/directives.html#figure
 */
export interface FigureOptions extends ImageOptions {
  /** Figure caption (single paragraph) */
  caption?: string;
  /** Figure legend (additional body elements) */
  legend?: string;
  /** Width of figure container ("image", length, or percentage) */
  figwidth?: string;
  /** CSS class names for the figure element */
  figclass?: string;
  /** Reference name for the figure element */
  figname?: string;
  /** Figure number (e.g., "1", "2.3") */
  figureNumber?: string;
}

/**
 * RST table directive options
 * @see https://docutils.sourceforge.io/docs/ref/rst/directives.html#table
 */
export interface TableOptions {
  /** Table caption */
  caption?: string;
  /** Table number (e.g., "1", "2") */
  tableNumber?: string;
  /** Horizontal alignment: left, center, or right */
  align?: 'left' | 'center' | 'right';
  /** Total table width (length or percentage) */
  width?: string;
  /** Column widths: "auto", "grid", or list of integers */
  widths?: string | number[];
  /** CSS class names (space-separated) */
  class?: string;
  /** Reference name for cross-references */
  name?: string;
  /** Whether first row is a header */
  hasHeader?: boolean;
}

/**
 * Table cell data
 */
export interface TableCell {
  /** Cell content */
  content: string;
  /** Column span */
  colspan?: number;
  /** Row span */
  rowspan?: number;
  /** Cell alignment */
  align?: 'left' | 'center' | 'right';
}

/**
 * Table row data
 */
export interface TableRow {
  /** Cells in the row */
  cells: TableCell[];
  /** Whether this is a header row */
  isHeader?: boolean;
}

/**
 * Table data structure
 */
export interface TableData {
  /** Table rows */
  rows: TableRow[];
  /** Table options */
  options: TableOptions;
}

/**
 * RST contents directive options
 * @see https://docutils.sourceforge.io/docs/ref/rst/directives.html#table-of-contents
 */
export interface ContentsOptions {
  /** TOC title (default: "Contents") */
  title?: string;
  /** Maximum heading levels to include */
  depth?: number;
  /** Only include subsections of current section */
  local?: boolean;
  /** Back-linking behavior: entry, top, or none */
  backlinks?: 'entry' | 'top' | 'none';
  /** CSS class names (space-separated) */
  class?: string;
}

/**
 * Custom directive parsed from rst_* style
 */
export interface CustomDirective {
  /** Directive name (e.g., "note", "warning", "need") */
  name: string;
  /** Optional directive argument */
  argument?: string;
  /** Directive options as key-value pairs */
  options: Map<string, string>;
  /** Directive body content */
  content: string;
}

/**
 * Parsed caption from Word
 */
export interface ParsedCaption {
  /** Caption type: "Figure", "Table", etc. */
  type: string;
  /** Caption number: "1", "2.1", etc. */
  number: string;
  /** Caption text after the number */
  text: string;
  /** Full original caption text */
  original: string;
}

/**
 * Extracted image data from Word document
 */
export interface ExtractedImage {
  /** Unique identifier */
  id: string;
  /** Generated filename (e.g., "image_001.png") */
  filename: string;
  /** Image binary data as base64 */
  base64Data: string;
  /** Image format: "png", "jpg", "gif", etc. */
  format: string;
  /** Original filename if available */
  originalName?: string;
  /** Original width in pixels */
  width?: number;
  /** Original height in pixels */
  height?: number;
  /** Alt text from Word */
  altText?: string;
}

/**
 * Document element types for conversion
 */
export type DocumentElementType =
  | 'paragraph'
  | 'heading'
  | 'list'
  | 'table'
  | 'image'
  | 'figure'
  | 'toc'
  | 'directive'
  | 'unknown';

/**
 * Base document element
 */
export interface DocumentElement {
  type: DocumentElementType;
  /** Original HTML content */
  html?: string;
  /** Word style name if applicable */
  style?: string;
}

/**
 * Heading element
 */
export interface HeadingElement extends DocumentElement {
  type: 'heading';
  /** Heading level (1-6) */
  level: number;
  /** Heading text */
  text: string;
}

/**
 * Paragraph element
 */
export interface ParagraphElement extends DocumentElement {
  type: 'paragraph';
  /** Paragraph content (may contain inline formatting) */
  content: string;
  /** Whether paragraph is a block quote */
  isBlockQuote?: boolean;
}

/**
 * List element
 */
export interface ListElement extends DocumentElement {
  type: 'list';
  /** List type: unordered or ordered */
  listType: 'unordered' | 'ordered';
  /** List items */
  items: ListItem[];
}

/**
 * List item
 */
export interface ListItem {
  /** Item content */
  content: string;
  /** Nested list if any */
  nestedList?: ListElement;
  /** Indent level (0 = top level, used during parsing) */
  indentLevel?: number;
}

/**
 * Image element
 */
export interface ImageElement extends DocumentElement {
  type: 'image';
  /** Image options */
  options: ImageOptions;
  /** Extracted image data */
  imageData?: ExtractedImage;
}

/**
 * Figure element
 */
export interface FigureElement extends DocumentElement {
  type: 'figure';
  /** Figure options */
  options: FigureOptions;
  /** Extracted image data */
  imageData?: ExtractedImage;
}

/**
 * Table element
 */
export interface TableElement extends DocumentElement {
  type: 'table';
  /** Table data */
  data: TableData;
}

/**
 * TOC element
 */
export interface TocElement extends DocumentElement {
  type: 'toc';
  /** Contents options */
  options: ContentsOptions;
}

/**
 * Custom directive element
 */
export interface DirectiveElement extends DocumentElement {
  type: 'directive';
  /** Parsed directive */
  directive: CustomDirective;
}

/**
 * Utility type for all document elements
 */
export type AnyDocumentElement =
  | HeadingElement
  | ParagraphElement
  | ListElement
  | ImageElement
  | FigureElement
  | TableElement
  | TocElement
  | DirectiveElement;
