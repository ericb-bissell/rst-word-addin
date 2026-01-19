/**
 * RST Word Add-in - HTML Parser
 * Parses Word's HTML output into structured document elements
 *
 * Word's getHtml() method returns HTML with specific patterns and styles
 * that need special handling for accurate RST conversion.
 */

import {
  AnyDocumentElement,
  HeadingElement,
  ParagraphElement,
  ListElement,
  ListItem,
  ImageElement,
  FigureElement,
  TableElement,
  TocElement,
  DirectiveElement,
  ExtractedImage,
  ImageOptions,
  FigureOptions,
} from './types';

import {
  parseHtmlTable,
  isTocElement,
  parseTocOptions,
  isRstDirectiveStyle,
  parseCustomDirective,
} from './directives';

import {
  findNearbyCaption,
  parseCaption,
  hasCaptionStyle,
} from '../utils/caption-parser';

/**
 * Result of parsing Word HTML
 */
export interface ParsedDocument {
  /** Document elements in order */
  elements: AnyDocumentElement[];
  /** Extracted images */
  images: ExtractedImage[];
  /** Document metadata */
  metadata: DocumentMetadata;
}

/**
 * Document metadata extracted from HTML
 */
export interface DocumentMetadata {
  /** Document title if found */
  title?: string;
  /** Author if found */
  author?: string;
  /** Language code (e.g., "en-US") */
  language?: string;
}

/**
 * Word heading style patterns
 */
const HEADING_PATTERNS = [
  // Word heading styles
  { pattern: /MsoHeading(\d)/i, levelExtractor: (m: RegExpMatchArray) => parseInt(m[1]) },
  { pattern: /Heading\s*(\d)/i, levelExtractor: (m: RegExpMatchArray) => parseInt(m[1]) },
  // HTML heading tags
  { pattern: /^h(\d)$/i, levelExtractor: (m: RegExpMatchArray) => parseInt(m[1]) },
];

// Word list style patterns (used for future list detection enhancements)
// const LIST_PATTERNS = {
//   ordered: /MsoListNumber|MsoListParagraph.*level.*numbering|list-style-type:\s*decimal/i,
//   unordered: /MsoListBullet|MsoListParagraph|list-style-type:\s*(disc|circle|square)/i,
// };

/**
 * Image counter for generating unique filenames
 */
let imageCounter = 0;

/**
 * Reset image counter (call before parsing a new document)
 */
export function resetImageCounter(): void {
  imageCounter = 0;
}

/**
 * Parse Word HTML into structured document elements
 *
 * @param html - HTML string from Word's getHtml()
 * @returns Parsed document with elements and images
 */
export function parseWordHtml(html: string): ParsedDocument {
  resetImageCounter();

  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  const elements: AnyDocumentElement[] = [];
  const images: ExtractedImage[] = [];
  const metadata = extractMetadata(doc);

  // Process body content
  const body = doc.body;
  if (!body) {
    return { elements, images, metadata };
  }

  // Get all top-level block elements
  const blockElements = getBlockElements(body);

  for (const element of blockElements) {
    const parsed = parseElement(element, images);
    if (parsed) {
      if (Array.isArray(parsed)) {
        elements.push(...parsed);
      } else {
        elements.push(parsed);
      }
    }
  }

  // Post-process: merge consecutive list items, handle figures with captions
  const processed = postProcessElements(elements);

  return { elements: processed, images, metadata };
}

/**
 * Extract metadata from HTML document
 */
function extractMetadata(doc: Document): DocumentMetadata {
  const metadata: DocumentMetadata = {};

  // Try to get title
  const titleEl = doc.querySelector('title');
  if (titleEl) {
    metadata.title = titleEl.textContent?.trim();
  }

  // Try to get author from meta tag
  const authorMeta = doc.querySelector('meta[name="author"]');
  if (authorMeta) {
    metadata.author = authorMeta.getAttribute('content') || undefined;
  }

  // Try to get language
  const htmlEl = doc.documentElement;
  if (htmlEl) {
    metadata.language = htmlEl.getAttribute('lang') || undefined;
  }

  return metadata;
}

/**
 * Get block-level elements from a container
 */
function getBlockElements(container: HTMLElement): HTMLElement[] {
  const elements: HTMLElement[] = [];
  const blockTags = ['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'DIV', 'TABLE', 'UL', 'OL', 'BLOCKQUOTE', 'FIGURE', 'NAV'];

  for (const child of Array.from(container.children)) {
    if (blockTags.includes(child.tagName)) {
      elements.push(child as HTMLElement);
    } else if (child.tagName === 'SPAN' || child.tagName === 'DIV') {
      // Some Word content is wrapped in spans/divs
      const nestedBlocks = getBlockElements(child as HTMLElement);
      if (nestedBlocks.length > 0) {
        elements.push(...nestedBlocks);
      } else if (child.textContent?.trim()) {
        elements.push(child as HTMLElement);
      }
    }
  }

  return elements;
}

/**
 * Parse a single HTML element into document element(s)
 */
function parseElement(
  element: HTMLElement,
  images: ExtractedImage[]
): AnyDocumentElement | AnyDocumentElement[] | null {
  const tagName = element.tagName.toUpperCase();
  const className = element.className || '';

  // Check for TOC first (can be div, nav, or have specific classes)
  if (isTocElement(element)) {
    return parseTocElement(element);
  }

  // Check for custom RST directive style
  const wordStyle = extractWordStyle(element);
  if (wordStyle && isRstDirectiveStyle(wordStyle)) {
    return parseDirectiveElement(element, wordStyle);
  }

  // Check for heading
  const headingLevel = detectHeadingLevel(element);
  if (headingLevel) {
    return parseHeadingElement(element, headingLevel);
  }

  // Check for table
  if (tagName === 'TABLE') {
    return parseTableElement(element);
  }

  // Check for list
  if (tagName === 'UL' || tagName === 'OL') {
    return parseListElement(element);
  }

  // Check for figure
  if (tagName === 'FIGURE') {
    return parseFigureElement(element, images);
  }

  // Check for image (standalone or with nearby caption)
  const img = element.querySelector('img') || (tagName === 'IMG' ? element as unknown as HTMLImageElement : null);
  if (img) {
    return parseImageOrFigure(element, img as HTMLImageElement, images);
  }

  // Check for blockquote
  if (tagName === 'BLOCKQUOTE' || className.toLowerCase().includes('quote')) {
    return parseBlockQuote(element);
  }

  // Check for caption (might have already been processed with its figure)
  if (hasCaptionStyle(element)) {
    // Standalone caption - return as paragraph with special note
    return {
      type: 'paragraph',
      content: element.textContent?.trim() || '',
      style: wordStyle || undefined,
      html: element.outerHTML,
    };
  }

  // Default: paragraph
  return parseParagraphElement(element);
}

/**
 * Extract Word style name from element
 */
function extractWordStyle(element: HTMLElement): string | null {
  const className = element.className || '';

  // Look for Mso* classes (e.g., MsoNormal, MsoHeading1)
  const msoMatch = className.match(/Mso(\w+)/);
  if (msoMatch) {
    return msoMatch[1];
  }

  // Look for style attribute with mso-style-name
  const style = element.getAttribute('style') || '';
  const styleNameMatch = style.match(/mso-style-name:\s*['"]?([^;'"]+)/i);
  if (styleNameMatch) {
    return styleNameMatch[1].trim();
  }

  // Check for data attribute
  const dataStyle = element.getAttribute('data-style');
  if (dataStyle) {
    return dataStyle;
  }

  // Check class list for rst_ prefixed style
  const classes = className.split(/\s+/);
  for (const cls of classes) {
    if (cls.toLowerCase().startsWith('rst_') || cls.startsWith('rst-')) {
      return cls.replace('rst-', 'rst_');
    }
  }

  return null;
}

/**
 * Detect heading level from element
 */
function detectHeadingLevel(element: HTMLElement): number | null {
  const tagName = element.tagName.toUpperCase();
  const className = element.className || '';

  // Check tag name first
  const tagMatch = tagName.match(/^H(\d)$/);
  if (tagMatch) {
    return parseInt(tagMatch[1]);
  }

  // Check class patterns
  for (const { pattern, levelExtractor } of HEADING_PATTERNS) {
    const match = className.match(pattern);
    if (match) {
      return levelExtractor(match);
    }
  }

  // Check style attribute for outline level
  const style = element.getAttribute('style') || '';
  const outlineMatch = style.match(/mso-outline-level:\s*(\d)/i);
  if (outlineMatch) {
    return parseInt(outlineMatch[1]);
  }

  return null;
}

/**
 * Parse heading element
 */
function parseHeadingElement(element: HTMLElement, level: number): HeadingElement {
  return {
    type: 'heading',
    level: Math.min(Math.max(level, 1), 6),
    text: getTextContent(element),
    html: element.outerHTML,
    style: extractWordStyle(element) || undefined,
  };
}

/**
 * Parse paragraph element
 */
function parseParagraphElement(element: HTMLElement): ParagraphElement | null {
  const content = getFormattedContent(element);

  // Skip empty paragraphs
  if (!content.trim()) {
    return null;
  }

  return {
    type: 'paragraph',
    content,
    html: element.outerHTML,
    style: extractWordStyle(element) || undefined,
  };
}

/**
 * Parse blockquote element
 */
function parseBlockQuote(element: HTMLElement): ParagraphElement {
  return {
    type: 'paragraph',
    content: getFormattedContent(element),
    isBlockQuote: true,
    html: element.outerHTML,
    style: extractWordStyle(element) || undefined,
  };
}

/**
 * Parse list element
 */
function parseListElement(element: HTMLElement): ListElement {
  const tagName = element.tagName.toUpperCase();
  const listType = tagName === 'OL' ? 'ordered' : 'unordered';

  const items: ListItem[] = [];
  const listItems = element.querySelectorAll(':scope > li');

  for (const li of Array.from(listItems)) {
    const item = parseListItem(li as HTMLElement);
    items.push(item);
  }

  return {
    type: 'list',
    listType,
    items,
    html: element.outerHTML,
  };
}

/**
 * Parse a list item
 */
function parseListItem(li: HTMLElement): ListItem {
  // Get content excluding nested lists
  let content = '';
  for (const child of Array.from(li.childNodes)) {
    if (child.nodeType === Node.TEXT_NODE) {
      content += child.textContent;
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      const el = child as HTMLElement;
      if (el.tagName !== 'UL' && el.tagName !== 'OL') {
        content += getFormattedContent(el);
      }
    }
  }

  const item: ListItem = {
    content: content.trim(),
  };

  // Check for nested list
  const nestedList = li.querySelector(':scope > ul, :scope > ol');
  if (nestedList) {
    item.nestedList = parseListElement(nestedList as HTMLElement);
  }

  return item;
}

/**
 * Parse table element
 */
function parseTableElement(element: HTMLElement): TableElement {
  const tableData = parseHtmlTable(element as HTMLTableElement);

  return {
    type: 'table',
    data: tableData,
    html: element.outerHTML,
  };
}

/**
 * Parse TOC element
 */
function parseTocElement(element: HTMLElement): TocElement {
  const options = parseTocOptions(element);

  return {
    type: 'toc',
    options,
    html: element.outerHTML,
  };
}

/**
 * Parse custom directive element
 */
function parseDirectiveElement(element: HTMLElement, styleName: string): DirectiveElement {
  const content = element.textContent || '';
  const directive = parseCustomDirective(styleName, content);

  return {
    type: 'directive',
    directive,
    html: element.outerHTML,
    style: styleName,
  };
}

/**
 * Check if an HTML element container is likely a figure
 * (e.g., centered image, image in a special container, etc.)
 */
function isLikelyFigure(element: HTMLElement): boolean {
  // Check for figure element
  if (element.tagName.toUpperCase() === 'FIGURE') {
    return true;
  }

  // Check for centered content (often figures)
  const style = element.getAttribute('style') || '';
  if (style.includes('text-align: center') || style.includes('text-align:center')) {
    return true;
  }

  // Check for figure-related classes
  const className = (element.className || '').toLowerCase();
  if (className.includes('figure') || className.includes('image-container')) {
    return true;
  }

  return false;
}

/**
 * Parse image or determine if it should be a figure
 */
function parseImageOrFigure(
  container: HTMLElement,
  img: HTMLImageElement,
  images: ExtractedImage[]
): ImageElement | FigureElement {
  const imageData = extractImageData(img);
  if (imageData) {
    images.push(imageData);
  }

  const imageOptions = parseImageOptionsFromElement(img, imageData);

  // Check if this should be a figure
  const captionElement = findNearbyCaption(container);

  // Convert to figure if there's a caption or if the container suggests figure treatment
  if (captionElement || isLikelyFigure(container)) {
    return createFigureElement(container, img, imageOptions, imageData, captionElement);
  }

  return {
    type: 'image',
    options: imageOptions,
    imageData,
    html: container.outerHTML,
  };
}

/**
 * Parse figure element
 */
function parseFigureElement(
  element: HTMLElement,
  images: ExtractedImage[]
): FigureElement | null {
  const img = element.querySelector('img');
  if (!img) {
    return null;
  }

  const imageData = extractImageData(img as HTMLImageElement);
  if (imageData) {
    images.push(imageData);
  }

  const imageOptions = parseImageOptionsFromElement(img as HTMLImageElement, imageData);
  const figcaption = element.querySelector('figcaption');

  return createFigureElement(element, img as HTMLImageElement, imageOptions, imageData, figcaption as HTMLElement | null);
}

/**
 * Create figure element from image and caption
 */
function createFigureElement(
  container: HTMLElement,
  _img: HTMLImageElement,
  imageOptions: ImageOptions,
  imageData: ExtractedImage | undefined,
  captionElement: HTMLElement | null
): FigureElement {
  const figureOptions: FigureOptions = { ...imageOptions };

  if (captionElement) {
    const captionText = captionElement.textContent?.trim() || '';
    const parsed = parseCaption(captionText);

    if (parsed) {
      figureOptions.caption = parsed.text;
      figureOptions.figureNumber = parsed.number;
      figureOptions.figname = `fig-${parsed.number.replace(/\./g, '-')}`;
    } else {
      figureOptions.caption = captionText;
    }
  }

  // Check for figure width
  const containerStyle = container.getAttribute('style') || '';
  const widthMatch = containerStyle.match(/width:\s*(\d+(?:\.\d+)?(?:px|%|em|cm|in|pt))/i);
  if (widthMatch) {
    figureOptions.figwidth = widthMatch[1];
  }

  return {
    type: 'figure',
    options: figureOptions,
    imageData,
    html: container.outerHTML,
  };
}

/**
 * Parse image options from img element
 */
function parseImageOptionsFromElement(
  img: HTMLImageElement,
  imageData: ExtractedImage | undefined
): ImageOptions {
  const src = img.getAttribute('src') || '';
  const alt = img.getAttribute('alt') || '';
  const style = img.getAttribute('style') || '';

  // Use extracted filename or original src
  const uri = imageData?.filename || src;

  const options: ImageOptions = {
    uri,
  };

  if (alt) {
    options.alt = alt;
  }

  // Parse dimensions from style or attributes
  const width = img.getAttribute('width') || extractStyleValue(style, 'width');
  const height = img.getAttribute('height') || extractStyleValue(style, 'height');

  if (width) {
    options.width = normalizeLength(width);
  }
  if (height) {
    options.height = normalizeLength(height);
  }

  // Check for alignment
  const align = detectImageAlignment(img);
  if (align) {
    options.align = align;
  }

  // Check for target (linked image)
  const parentLink = img.closest('a');
  if (parentLink) {
    const href = parentLink.getAttribute('href');
    if (href && !href.startsWith('#')) {
      options.target = href;
    }
  }

  return options;
}

/**
 * Extract image data from img element
 */
function extractImageData(img: HTMLImageElement): ExtractedImage | undefined {
  const src = img.getAttribute('src') || '';

  // Handle base64 data URLs
  if (src.startsWith('data:image/')) {
    const match = src.match(/^data:image\/(\w+);base64,(.+)$/);
    if (match) {
      const format = match[1] === 'jpeg' ? 'jpg' : match[1];
      const base64Data = match[2];

      imageCounter++;
      const filename = `images/image_${String(imageCounter).padStart(3, '0')}.${format}`;

      return {
        id: `img-${imageCounter}`,
        filename,
        base64Data,
        format,
        altText: img.getAttribute('alt') || undefined,
        width: img.naturalWidth || undefined,
        height: img.naturalHeight || undefined,
      };
    }
  }

  // Handle external URLs (we'll reference them directly)
  if (src.startsWith('http://') || src.startsWith('https://')) {
    return undefined; // Will use URL directly
  }

  // Handle blob URLs or other formats
  if (src) {
    imageCounter++;
    const extension = getExtensionFromSrc(src);
    const filename = `images/image_${String(imageCounter).padStart(3, '0')}.${extension}`;

    return {
      id: `img-${imageCounter}`,
      filename,
      base64Data: '', // Will need to be fetched
      format: extension,
      altText: img.getAttribute('alt') || undefined,
    };
  }

  return undefined;
}

/**
 * Get file extension from src
 */
function getExtensionFromSrc(src: string): string {
  const match = src.match(/\.(\w+)(?:\?.*)?$/);
  if (match) {
    const ext = match[1].toLowerCase();
    return ext === 'jpeg' ? 'jpg' : ext;
  }
  return 'png'; // Default
}

/**
 * Detect image alignment from element
 */
function detectImageAlignment(img: HTMLImageElement): ImageOptions['align'] | undefined {
  const style = img.getAttribute('style') || '';
  const className = img.className || '';

  // Check for float
  if (style.includes('float: left') || style.includes('float:left')) {
    return 'left';
  }
  if (style.includes('float: right') || style.includes('float:right')) {
    return 'right';
  }

  // Check for text-align on parent
  const parent = img.parentElement;
  if (parent) {
    const parentStyle = parent.getAttribute('style') || '';
    if (parentStyle.includes('text-align: center') || parentStyle.includes('text-align:center')) {
      return 'center';
    }
    if (parentStyle.includes('text-align: left')) {
      return 'left';
    }
    if (parentStyle.includes('text-align: right')) {
      return 'right';
    }
  }

  // Check for class-based alignment
  const alignClasses = ['left', 'center', 'right', 'middle', 'top', 'bottom'];
  for (const align of alignClasses) {
    if (className.includes(align)) {
      return align as ImageOptions['align'];
    }
  }

  return undefined;
}

/**
 * Extract style value from inline style string
 */
function extractStyleValue(style: string, property: string): string | null {
  const regex = new RegExp(`${property}:\\s*([^;]+)`, 'i');
  const match = style.match(regex);
  return match ? match[1].trim() : null;
}

/**
 * Normalize length value (add px if no unit)
 */
function normalizeLength(value: string): string {
  const trimmed = value.trim();
  if (/^\d+$/.test(trimmed)) {
    return `${trimmed}px`;
  }
  return trimmed;
}

/**
 * Get text content with basic formatting preserved
 */
function getTextContent(element: HTMLElement): string {
  return element.textContent?.trim() || '';
}

/**
 * Get formatted content preserving inline RST markup
 */
function getFormattedContent(element: HTMLElement): string {
  let result = '';

  for (const node of Array.from(element.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE) {
      result += node.textContent;
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as HTMLElement;
      const tag = el.tagName.toUpperCase();

      switch (tag) {
        case 'STRONG':
        case 'B':
          result += `**${getFormattedContent(el)}**`;
          break;
        case 'EM':
        case 'I':
          result += `*${getFormattedContent(el)}*`;
          break;
        case 'CODE':
        case 'TT':
          result += `\`\`${getFormattedContent(el)}\`\``;
          break;
        case 'SUB':
          result += `:sub:\`${getFormattedContent(el)}\``;
          break;
        case 'SUP':
          result += `:sup:\`${getFormattedContent(el)}\``;
          break;
        case 'A':
          const href = el.getAttribute('href');
          const text = getFormattedContent(el);
          if (href) {
            if (href.startsWith('#')) {
              // Internal link
              result += `:ref:\`${text} <${href.substring(1)}>\``;
            } else {
              // External link
              result += `\`${text} <${href}>\`_`;
            }
          } else {
            result += text;
          }
          break;
        case 'BR':
          result += '\n';
          break;
        case 'SPAN':
          // Check for special formatting
          const style = el.getAttribute('style') || '';
          const className = el.className || '';
          const content = getFormattedContent(el);

          if (style.includes('font-weight') && (style.includes('bold') || style.includes('700'))) {
            result += `**${content}**`;
          } else if (style.includes('font-style') && style.includes('italic')) {
            result += `*${content}*`;
          } else if (style.includes('text-decoration') && style.includes('underline')) {
            // RST doesn't have underline, use emphasis
            result += `*${content}*`;
          } else if (className.includes('strike') || style.includes('line-through')) {
            // Strikethrough - no standard RST, could use custom role
            result += content;
          } else {
            result += content;
          }
          break;
        default:
          result += getFormattedContent(el);
      }
    }
  }

  return result;
}

/**
 * Post-process elements to handle special cases
 */
function postProcessElements(elements: AnyDocumentElement[]): AnyDocumentElement[] {
  const result: AnyDocumentElement[] = [];

  for (let i = 0; i < elements.length; i++) {
    const current = elements[i];

    // Skip caption paragraphs that were already processed with figures
    if (current.type === 'paragraph') {
      const para = current as ParagraphElement;
      // Check if this caption was already used by a previous figure
      if (hasCaptionStyle({ className: para.style || '' } as HTMLElement)) {
        const prev = result[result.length - 1];
        if (prev && prev.type === 'figure') {
          // Skip this caption, it was already attached
          continue;
        }
      }
    }

    // Merge consecutive directives with the same style
    if (current.type === 'directive') {
      const directive = current as DirectiveElement;
      const last = result[result.length - 1];

      if (last && last.type === 'directive') {
        const lastDirective = last as DirectiveElement;
        if (lastDirective.style === directive.style) {
          // Merge content
          lastDirective.directive.content += '\n\n' + directive.directive.content;
          continue;
        }
      }
    }

    result.push(current);
  }

  return result;
}
