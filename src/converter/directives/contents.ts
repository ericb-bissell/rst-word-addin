/**
 * RST Word Add-in - Contents Directive Generator
 * Generates RST `.. contents::` directives for table of contents
 *
 * @see https://docutils.sourceforge.io/docs/ref/rst/directives.html#table-of-contents
 */

import { ContentsOptions } from '../types';

/**
 * Default indentation for directive options
 */
const INDENT = '   ';

/**
 * Generate an RST contents directive
 *
 * @param options - Contents directive options
 * @returns RST contents directive string
 *
 * @example
 * ```typescript
 * const rst = generateContentsDirective({
 *   title: 'Table of Contents',
 *   depth: 3,
 *   backlinks: 'entry'
 * });
 * // Returns:
 * // .. contents:: Table of Contents
 * //    :depth: 3
 * //    :backlinks: entry
 * ```
 */
export function generateContentsDirective(options: ContentsOptions = {}): string {
  const lines: string[] = [];

  // Directive declaration with optional title
  if (options.title) {
    lines.push(`.. contents:: ${options.title}`);
  } else {
    lines.push('.. contents::');
  }

  // Add options
  if (options.depth !== undefined && options.depth > 0) {
    lines.push(`${INDENT}:depth: ${options.depth}`);
  }

  if (options.local) {
    lines.push(`${INDENT}:local:`);
  }

  if (options.backlinks) {
    lines.push(`${INDENT}:backlinks: ${options.backlinks}`);
  }

  if (options.class) {
    lines.push(`${INDENT}:class: ${options.class}`);
  }

  return lines.join('\n');
}

/**
 * Detect if content represents a Table of Contents from Word
 *
 * Word TOC can be identified by:
 * - TOC field codes
 * - Paragraph styles like "TOC Heading", "TOC 1", "TOC 2"
 * - Specific HTML patterns
 *
 * @param element - HTML element to check
 * @returns True if this appears to be a TOC
 */
export function isTocElement(element: HTMLElement): boolean {
  // Check for TOC-related classes
  const className = element.className?.toLowerCase() || '';
  if (
    className.includes('toc') ||
    className.includes('tableofcontents') ||
    className.includes('table-of-contents')
  ) {
    return true;
  }

  // Check for Word TOC field code in data attributes
  const fieldCode = element.getAttribute('data-field-code');
  if (fieldCode && fieldCode.toLowerCase().includes('toc')) {
    return true;
  }

  // Check for TOC style attributes
  const style = element.getAttribute('style') || '';
  if (style.includes('mso-toc') || style.includes('TOC')) {
    return true;
  }

  // Check for typical TOC structure (heading + nested list of links)
  if (element.tagName === 'DIV' || element.tagName === 'NAV') {
    const hasHeading = element.querySelector('h1, h2, h3, h4, [class*="heading"]');
    const hasLinks = element.querySelectorAll('a').length > 3;
    const hasNestedStructure = element.querySelectorAll('ul, ol').length > 0;

    if (hasHeading && hasLinks && hasNestedStructure) {
      // Check if links appear to be internal document links
      const links = element.querySelectorAll('a');
      let internalLinkCount = 0;
      links.forEach((link) => {
        const href = link.getAttribute('href');
        if (href && (href.startsWith('#') || href.startsWith('_'))) {
          internalLinkCount++;
        }
      });

      if (internalLinkCount > links.length / 2) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Parse TOC options from Word TOC field or HTML
 *
 * @param element - TOC HTML element
 * @returns Parsed contents options
 */
export function parseTocOptions(element: HTMLElement): ContentsOptions {
  const options: ContentsOptions = {};

  // Try to extract title from heading
  const heading = element.querySelector('h1, h2, h3, h4, [class*="heading"]');
  if (heading) {
    const text = heading.textContent?.trim();
    if (text && text.toLowerCase() !== 'contents') {
      options.title = text;
    }
  }

  // Try to determine depth from structure
  const maxDepth = calculateTocDepth(element);
  if (maxDepth > 0 && maxDepth < 10) {
    options.depth = maxDepth;
  }

  // Check for field code that might contain options
  const fieldCode = element.getAttribute('data-field-code');
  if (fieldCode) {
    const parsedOptions = parseWordTocFieldCode(fieldCode);
    Object.assign(options, parsedOptions);
  }

  return options;
}

/**
 * Calculate the depth of a TOC from its HTML structure
 *
 * @param element - TOC element
 * @returns Maximum depth found
 */
function calculateTocDepth(element: HTMLElement): number {
  let maxDepth = 0;

  // Count nested list levels
  const countDepth = (el: Element, currentDepth: number): void => {
    maxDepth = Math.max(maxDepth, currentDepth);

    const nestedLists = el.querySelectorAll(':scope > li > ul, :scope > li > ol');
    nestedLists.forEach((list) => {
      countDepth(list, currentDepth + 1);
    });
  };

  const topLevelLists = element.querySelectorAll(':scope > ul, :scope > ol');
  topLevelLists.forEach((list) => {
    countDepth(list, 1);
  });

  return maxDepth;
}

/**
 * Parse Word TOC field code for options
 *
 * Word TOC field codes look like: TOC \o "1-3" \h \z \u
 * - \o "1-3" = outline levels 1-3
 * - \h = hyperlinks
 * - \z = hide tab leaders and page numbers in Web Layout
 * - \u = use paragraph outline level
 *
 * @param fieldCode - Word field code string
 * @returns Parsed options
 */
function parseWordTocFieldCode(fieldCode: string): ContentsOptions {
  const options: ContentsOptions = {};

  // Extract outline level range
  const outlineMatch = fieldCode.match(/\\o\s*"(\d+)-(\d+)"/i);
  if (outlineMatch) {
    const startLevel = parseInt(outlineMatch[1]);
    const endLevel = parseInt(outlineMatch[2]);
    options.depth = endLevel - startLevel + 1;
  }

  // Check for single level specification
  const levelMatch = fieldCode.match(/\\l\s*"?(\d+)"?/i);
  if (levelMatch) {
    options.depth = parseInt(levelMatch[1]);
  }

  return options;
}

/**
 * Determine if backlinks should be enabled based on document context
 *
 * @param hasPageNumbers - Whether original TOC has page numbers
 * @param isWebOutput - Whether output is for web/HTML
 * @returns Recommended backlinks setting
 */
export function recommendBacklinks(
  hasPageNumbers: boolean,
  isWebOutput: boolean
): ContentsOptions['backlinks'] {
  // For web output without page numbers, backlinks are useful
  if (isWebOutput && !hasPageNumbers) {
    return 'entry';
  }

  // For documents with page numbers, backlinks might be redundant
  if (hasPageNumbers) {
    return 'none';
  }

  // Default to entry for most cases
  return 'entry';
}

/**
 * Check if element contains Word TOC field markers
 *
 * @param html - HTML string to check
 * @returns True if contains TOC field
 */
export function containsTocField(html: string): boolean {
  // Check for common TOC field patterns
  const patterns = [
    /TOC\s*\\[ohzu]/i,
    /HYPERLINK\s*\\l\s*"_Toc/i,
    /class="?MsoToc/i,
    /style="?mso-toc/i,
    /<!\-\-\[if\s+supportFields\]>.*TOC.*<!\[endif\]\-\->/is,
  ];

  return patterns.some((pattern) => pattern.test(html));
}

/**
 * Extract TOC title from various HTML patterns
 *
 * @param element - Element that might contain TOC
 * @returns Extracted title or undefined
 */
export function extractTocTitle(element: HTMLElement): string | undefined {
  // Check for explicit heading
  const headings = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'];

  for (const tag of headings) {
    const heading = element.querySelector(tag);
    if (heading) {
      const text = heading.textContent?.trim();
      if (text) {
        return text;
      }
    }
  }

  // Check for paragraph with TOC heading style
  const tocHeading = element.querySelector('[class*="TocHeading"], [class*="TOCHeading"]');
  if (tocHeading) {
    return tocHeading.textContent?.trim();
  }

  // Check first child if it looks like a title
  const firstChild = element.firstElementChild;
  if (firstChild) {
    const text = firstChild.textContent?.trim();
    if (text && text.length < 50) {
      // Check if it's not a TOC entry (usually has numbers or specific patterns)
      if (!/^\d+(\.\d+)*\s/.test(text) && !text.includes('...')) {
        return text;
      }
    }
  }

  return undefined;
}
