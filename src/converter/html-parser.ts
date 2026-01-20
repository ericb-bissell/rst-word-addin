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
  // Content block tags (not wrapper divs)
  const contentBlockTags = ['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'TABLE', 'UL', 'OL', 'BLOCKQUOTE', 'FIGURE', 'NAV'];
  // Wrapper tags that we should recurse into
  const wrapperTags = ['DIV', 'SPAN', 'SECTION', 'ARTICLE', 'MAIN'];

  for (const child of Array.from(container.children)) {
    const tagName = child.tagName.toUpperCase();

    if (contentBlockTags.includes(tagName)) {
      // Direct content block - add it
      elements.push(child as HTMLElement);
    } else if (wrapperTags.includes(tagName)) {
      // Wrapper element - recurse into it to find content blocks
      const nestedBlocks = getBlockElements(child as HTMLElement);
      if (nestedBlocks.length > 0) {
        elements.push(...nestedBlocks);
      } else if (child.textContent?.trim()) {
        // Wrapper has text content but no block children - treat as paragraph
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
    // Check if this is a layout table (used for positioning images) vs content table
    if (isLayoutTable(element as HTMLTableElement)) {
      // Extract images from layout table
      const layoutImages = extractImagesFromLayoutTable(element as HTMLTableElement, images);
      if (layoutImages.length > 0) {
        return layoutImages;
      }
      // No images found, skip this layout table
      return null;
    }
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

  // Check for images (standalone or with nearby caption)
  // First check if there's a layout table inside with images
  const nestedLayoutTable = element.querySelector('table') as HTMLTableElement | null;
  if (nestedLayoutTable && isLayoutTable(nestedLayoutTable)) {
    const layoutImages = extractImagesFromLayoutTable(nestedLayoutTable, images);
    if (layoutImages.length > 0) {
      return layoutImages;
    }
  }

  // Check for images directly in the element
  const allImgs = element.querySelectorAll('img');
  if (allImgs.length > 1) {
    // Multiple images - return as array
    const imageElements: AnyDocumentElement[] = [];
    for (const imgEl of Array.from(allImgs)) {
      const imgHtml = imgEl as HTMLImageElement;
      const imageData = extractImageData(imgHtml);
      if (imageData) {
        images.push(imageData);
      }
      const imageOptions = parseImageOptionsFromElement(imgHtml, imageData);
      imageElements.push({
        type: 'image',
        options: imageOptions,
        imageData,
        html: imgHtml.outerHTML,
      });
    }
    return imageElements;
  } else if (allImgs.length === 1) {
    return parseImageOrFigure(element, allImgs[0] as HTMLImageElement, images);
  }

  const img = tagName === 'IMG' ? element as unknown as HTMLImageElement : null;
  if (img) {
    return parseImageOrFigure(element, img, images);
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

  // Check for Word list paragraph (MsoListParagraph variants)
  if (isWordListParagraph(element)) {
    return parseWordListItem(element);
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
 * Check if element is a Word list paragraph
 */
function isWordListParagraph(element: HTMLElement): boolean {
  const className = element.className || '';
  return /MsoListParagraph/i.test(className);
}

/**
 * Extract indent level from Word's margin-left style
 * Word uses .5in increments: .5in = level 0, 1.0in = level 1, 1.5in = level 2, etc.
 */
function extractIndentLevel(element: HTMLElement): number {
  const style = element.getAttribute('style') || '';
  const marginMatch = style.match(/margin-left:\s*([\d.]+)(in|pt|cm)/i);

  if (marginMatch) {
    let inches = parseFloat(marginMatch[1]);
    const unit = marginMatch[2].toLowerCase();

    // Convert to inches
    if (unit === 'pt') {
      inches = inches / 72;
    } else if (unit === 'cm') {
      inches = inches / 2.54;
    }

    // .5in = level 0, 1.0in = level 1, etc.
    return Math.max(0, Math.round((inches - 0.5) / 0.5));
  }

  return 0;
}

/**
 * Detect if list item is ordered by examining the marker
 */
function detectListType(element: HTMLElement): 'ordered' | 'unordered' {
  // Get the raw text content to check for number/letter markers
  const fullText = element.textContent || '';

  // Check for numbered list markers at the start: "1.", "2.", "a.", "b.", "i.", "ii.", etc.
  // Word formats these as plain text before the content
  if (/^\s*(\d+|[a-z]|[ivxlcdm]+)[.\)]\s/i.test(fullText)) {
    return 'ordered';
  }

  // Check for bullet characters (Symbol/Wingdings fonts produce these)
  // ·, •, ◦, o, §, ▪, etc. indicate unordered
  if (/^[\s·•◦▪▸►§o\-\*]+/.test(fullText)) {
    return 'unordered';
  }

  // Default to unordered
  return 'unordered';
}

/**
 * Parse Word list item from MsoListParagraph element
 */
function parseWordListItem(element: HTMLElement): ListElement {
  // Extract indent level and list type
  const indentLevel = extractIndentLevel(element);
  const listType = detectListType(element);

  // Get the text content, removing the bullet/number span
  let content = '';

  // Word puts the bullet in a span with font-family:Symbol or Wingdings
  // The actual content follows after
  for (const node of Array.from(element.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE) {
      content += node.textContent;
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as HTMLElement;
      const style = el.getAttribute('style') || '';
      const fontFamily = style.toLowerCase();

      // Skip bullet/number spans (Symbol, Wingdings, Courier New for 'o' bullets)
      if (fontFamily.includes('symbol') ||
          fontFamily.includes('wingdings') ||
          fontFamily.includes('courier')) {
        continue;
      }

      // Check nested spans for bullet markers
      if (el.tagName === 'SPAN') {
        const innerStyle = el.querySelector('[style*="Symbol"], [style*="Wingdings"], [style*="Courier"]');
        if (innerStyle) {
          // This span contains the bullet, extract only the text after it
          const textParts: string[] = [];
          for (const child of Array.from(el.childNodes)) {
            if (child.nodeType === Node.TEXT_NODE) {
              textParts.push(child.textContent || '');
            } else if (child.nodeType === Node.ELEMENT_NODE) {
              const childEl = child as HTMLElement;
              const childStyle = childEl.getAttribute('style') || '';
              if (!childStyle.includes('Symbol') &&
                  !childStyle.includes('Wingdings') &&
                  !childStyle.includes('Courier')) {
                textParts.push(childEl.textContent || '');
              }
            }
          }
          content += textParts.join('');
        } else {
          content += getFormattedContent(el);
        }
      } else {
        content += getFormattedContent(el);
      }
    }
  }

  // Clean up the content - remove bullet characters, numbers, letters and extra whitespace
  content = content
    .replace(/^[\s·•◦▪▸►§o\-\*]+/, '')           // Remove leading bullets
    .replace(/^\s*(\d+|[a-z]|[ivxlcdm]+)[.\)]\s*/i, '') // Remove leading numbers/letters/roman
    .replace(/\s+/g, ' ')                          // Normalize whitespace
    .trim();

  // Return as a single-item list with indent level
  return {
    type: 'list',
    listType,
    items: [{ content, indentLevel }],
    html: element.outerHTML,
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
 * Check if a table is a layout table (used for positioning images) vs content table
 * Word uses layout tables with cellpadding=0 and cellspacing=0 to position images
 * Content tables have the MsoTableGrid class
 */
function isLayoutTable(table: HTMLTableElement): boolean {
  const className = table.className || '';

  // Content tables have MsoTableGrid class
  if (className.includes('MsoTableGrid')) {
    return false;
  }

  // Check for layout table indicators
  const cellPadding = table.getAttribute('cellpadding');
  const cellSpacing = table.getAttribute('cellspacing');
  const hasImages = table.querySelectorAll('img').length > 0;

  // Layout tables typically have cellpadding=0 and cellspacing=0 and contain images
  const isLayoutStyle = cellPadding === '0' || cellSpacing === '0';

  return hasImages && isLayoutStyle;
}

/**
 * Extract all images from a layout table
 * Layout tables are used by Word to position multiple images
 */
function extractImagesFromLayoutTable(
  table: HTMLTableElement,
  images: ExtractedImage[]
): ImageElement[] {
  const result: ImageElement[] = [];
  const imgElements = table.querySelectorAll('img');

  for (const imgEl of Array.from(imgElements)) {
    const img = imgEl as HTMLImageElement;
    const imageData = extractImageData(img);

    if (imageData) {
      images.push(imageData);
    }

    const imageOptions = parseImageOptionsFromElement(img, imageData);

    result.push({
      type: 'image',
      options: imageOptions,
      imageData,
      html: img.outerHTML,
    });
  }

  return result;
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

  // Always set a unique figure name based on image ID
  const figureId = imageData?.id || `img-${imageCounter}`;
  figureOptions.figname = figureId;

  if (captionElement) {
    const captionText = captionElement.textContent?.trim() || '';
    const parsed = parseCaption(captionText);

    if (parsed) {
      figureOptions.caption = parsed.text;
      figureOptions.figureNumber = parsed.number;
      // Use parsed number for name if available
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
 * Normalize whitespace - collapse multiple spaces/newlines into single space
 */
function normalizeWhitespace(text: string): string {
  return text.replace(/\s+/g, ' ');
}

/**
 * Get text content with basic formatting preserved
 */
function getTextContent(element: HTMLElement): string {
  return normalizeWhitespace(element.textContent || '').trim();
}

/**
 * Get formatted content preserving inline RST markup
 */
function getFormattedContent(element: HTMLElement): string {
  let result = '';

  for (const node of Array.from(element.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE) {
      // Normalize whitespace in text nodes (collapse newlines/spaces)
      result += normalizeWhitespace(node.textContent || '');
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
 * Helper to add an item to a list at the correct nesting level
 *
 * @param rootList - The root list element
 * @param newItem - The new item to add
 * @param targetIndent - The indent level where item should be added
 * @param itemListType - The list type (ordered/unordered) of the new item
 */
function addItemToList(
  rootList: ListElement,
  newItem: ListItem,
  targetIndent: number,
  itemListType: 'ordered' | 'unordered'
): void {
  // Clean up the indentLevel from the item (it's now structural)
  delete newItem.indentLevel;

  // Find the correct place to add the item
  if (targetIndent === 0) {
    // Add to root level
    // If list types match, add as sibling; otherwise start nested list
    if (rootList.listType === itemListType) {
      rootList.items.push(newItem);
    } else {
      // Different list type at same level - add to last item as nested
      const lastItem = rootList.items[rootList.items.length - 1];
      if (lastItem) {
        if (!lastItem.nestedList) {
          lastItem.nestedList = {
            type: 'list',
            listType: itemListType,
            items: [],
            html: '',
          };
        }
        lastItem.nestedList.items.push(newItem);
      }
    }
    return;
  }

  // Need to go deeper - find the last item and recurse
  let currentList = rootList;
  let currentIndent = 0;

  while (currentIndent < targetIndent) {
    const lastItem = currentList.items[currentList.items.length - 1];
    if (!lastItem) {
      // No item to nest under, add to current level
      currentList.items.push(newItem);
      return;
    }

    if (currentIndent + 1 === targetIndent) {
      // This is where we need to add the item
      if (!lastItem.nestedList) {
        lastItem.nestedList = {
          type: 'list',
          listType: itemListType,
          items: [],
          html: '',
        };
      }
      lastItem.nestedList.items.push(newItem);
      return;
    }

    // Go deeper
    if (!lastItem.nestedList) {
      // Create intermediate nested list
      lastItem.nestedList = {
        type: 'list',
        listType: itemListType,
        items: [],
        html: '',
      };
    }
    currentList = lastItem.nestedList;
    currentIndent++;
  }

  // Add at current level
  currentList.items.push(newItem);
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
      const styleLC = (para.style || '').toLowerCase();
      const isCaptionStyle = styleLC.includes('caption') || styleLC.includes('figcaption');
      if (isCaptionStyle) {
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

    // Build nested list structures from flat list items with indent levels
    if (current.type === 'list') {
      const list = current as ListElement;
      const newItem = list.items[0]; // Each parsed list has exactly one item
      const newIndent = newItem.indentLevel ?? 0;
      const last = result[result.length - 1];

      if (last && last.type === 'list') {
        // Add to existing list structure
        const lastList = last as ListElement;
        addItemToList(lastList, newItem, newIndent, list.listType);
        continue;
      } else {
        // Start a new top-level list
        // Clear indentLevel from item since it's now structural
        delete newItem.indentLevel;
        result.push(list);
        continue;
      }
    }

    // Associate caption paragraphs with following tables
    if (current.type === 'table') {
      const table = current as TableElement;
      const prev = result[result.length - 1];

      if (prev && prev.type === 'paragraph') {
        const para = prev as ParagraphElement;
        const styleLC = (para.style || '').toLowerCase();
        const contentLC = para.content.toLowerCase();

        // Check if previous paragraph is a table caption
        if (styleLC.includes('caption') || contentLC.startsWith('table ')) {
          // Parse the caption
          const captionText = para.content.trim();
          const parsed = parseCaption(captionText);

          if (parsed && parsed.type === 'Table') {
            table.data.options.caption = parsed.text;
            table.data.options.tableNumber = parsed.number;
            table.data.options.name = `table-${parsed.number.replace(/\./g, '-')}`;
          } else if (styleLC.includes('caption')) {
            // MsoCaption style but couldn't parse - use full text
            table.data.options.caption = captionText;
          }

          // Remove the caption paragraph from results (it's now part of the table)
          if (table.data.options.caption) {
            result.pop();
          }
        }
      }
    }

    result.push(current);
  }

  return result;
}
