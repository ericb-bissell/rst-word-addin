/**
 * RST Word Add-in - Caption Parser
 * Parses Word captions for figures, tables, and other labeled elements
 *
 * Word captions typically follow patterns like:
 * - "Figure 1: Description"
 * - "Table 2.1: Title"
 * - "Listing 3 - Code example"
 */

import { ParsedCaption } from '../converter/types';

/**
 * Common caption type patterns
 */
const CAPTION_PATTERNS: Array<{
  type: string;
  pattern: RegExp;
}> = [
  // Figure patterns
  { type: 'Figure', pattern: /^(Figure|Fig\.?)\s+/i },

  // Table patterns
  { type: 'Table', pattern: /^(Table|Tbl\.?)\s+/i },

  // Listing/Code patterns
  { type: 'Listing', pattern: /^(Listing|List\.?|Code)\s+/i },

  // Equation patterns
  { type: 'Equation', pattern: /^(Equation|Eq\.?)\s+/i },

  // Example patterns
  { type: 'Example', pattern: /^(Example|Ex\.?)\s+/i },

  // Chart patterns
  { type: 'Chart', pattern: /^(Chart)\s+/i },

  // Diagram patterns
  { type: 'Diagram', pattern: /^(Diagram)\s+/i },

  // Generic numbered pattern (e.g., "1: Description")
  { type: 'Item', pattern: /^(\d+(?:\.\d+)*)\s*[:.]\s*/ },
];

/**
 * Parse a caption string to extract type, number, and text
 *
 * @param caption - Raw caption string
 * @returns Parsed caption components or null if not a valid caption
 *
 * @example
 * ```typescript
 * parseCaption("Figure 1: System Architecture")
 * // Returns: { type: "Figure", number: "1", text: "System Architecture", original: "..." }
 *
 * parseCaption("Table 2.1 - User Data")
 * // Returns: { type: "Table", number: "2.1", text: "User Data", original: "..." }
 * ```
 */
export function parseCaption(caption: string): ParsedCaption | null {
  const trimmed = caption.trim();

  if (!trimmed) {
    return null;
  }

  // Try each pattern
  for (const { type, pattern } of CAPTION_PATTERNS) {
    const match = trimmed.match(pattern);

    if (match) {
      // Remove the matched prefix
      const remainder = trimmed.substring(match[0].length);

      // Extract number
      const numberMatch = remainder.match(/^(\d+(?:\.\d+)*)/);
      let number = '';
      let textStart = 0;

      if (numberMatch) {
        number = numberMatch[1];
        textStart = numberMatch[0].length;
      }

      // Extract text after separator (: or -)
      let text = remainder.substring(textStart);
      const separatorMatch = text.match(/^\s*[:.–—-]\s*/);
      if (separatorMatch) {
        text = text.substring(separatorMatch[0].length);
      }

      text = text.trim();

      // For generic numbered pattern, the number is in the match itself
      if (type === 'Item' && match[1]) {
        return {
          type: 'Item',
          number: match[1],
          text: text,
          original: trimmed,
        };
      }

      return {
        type,
        number,
        text,
        original: trimmed,
      };
    }
  }

  // Check if it starts with a number (might be a simple numbered caption)
  const simpleNumberMatch = trimmed.match(/^(\d+(?:\.\d+)*)\s*[:.–—-]\s*(.*)$/);
  if (simpleNumberMatch) {
    return {
      type: 'Item',
      number: simpleNumberMatch[1],
      text: simpleNumberMatch[2].trim(),
      original: trimmed,
    };
  }

  return null;
}

/**
 * Check if a string looks like a caption
 *
 * @param text - Text to check
 * @returns True if text appears to be a caption
 */
export function looksLikeCaption(text: string): boolean {
  const trimmed = text.trim();

  // Check for caption patterns
  for (const { pattern } of CAPTION_PATTERNS) {
    if (pattern.test(trimmed)) {
      return true;
    }
  }

  // Check for simple numbered pattern
  if (/^\d+(?:\.\d+)*\s*[:.–—-]/.test(trimmed)) {
    return true;
  }

  return false;
}

/**
 * Detect caption type from text
 *
 * @param text - Text to analyze
 * @returns Detected caption type or null
 */
export function detectCaptionType(text: string): string | null {
  const trimmed = text.trim();

  for (const { type, pattern } of CAPTION_PATTERNS) {
    if (pattern.test(trimmed)) {
      return type;
    }
  }

  return null;
}

/**
 * Extract just the caption number
 *
 * @param caption - Caption text
 * @returns Extracted number or null
 */
export function extractCaptionNumber(caption: string): string | null {
  const parsed = parseCaption(caption);
  return parsed?.number || null;
}

/**
 * Extract just the caption text (without type and number)
 *
 * @param caption - Caption text
 * @returns Extracted description text
 */
export function extractCaptionText(caption: string): string {
  const parsed = parseCaption(caption);
  return parsed?.text || caption.trim();
}

/**
 * Format a caption for RST output
 *
 * @param parsed - Parsed caption
 * @param includeNumber - Whether to include the number
 * @returns Formatted caption string
 */
export function formatCaptionForRst(
  parsed: ParsedCaption,
  includeNumber: boolean = true
): string {
  if (includeNumber && parsed.number) {
    return `${parsed.type} ${parsed.number}: ${parsed.text}`;
  }
  return parsed.text;
}

/**
 * Generate a reference label from caption
 *
 * @param parsed - Parsed caption
 * @returns Reference label suitable for RST
 */
export function generateRefLabel(parsed: ParsedCaption): string {
  const prefix = parsed.type.toLowerCase().substring(0, 3);
  const number = parsed.number.replace(/\./g, '-');

  if (number) {
    return `${prefix}-${number}`;
  }

  // Generate from text
  const slug = parsed.text
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .substring(0, 40);

  return `${prefix}-${slug || 'unnamed'}`;
}

/**
 * Check if an element has Word caption styling
 *
 * @param element - HTML element to check
 * @returns True if element has caption styling
 */
export function hasCaptionStyle(element: HTMLElement): boolean {
  const className = element.className?.toLowerCase() || '';
  const style = element.getAttribute('style')?.toLowerCase() || '';

  // Check for Word caption class
  if (
    className.includes('caption') ||
    className.includes('msocaption') ||
    className.includes('figcaption')
  ) {
    return true;
  }

  // Check for caption style
  if (
    style.includes('caption') ||
    style.includes('mso-caption')
  ) {
    return true;
  }

  // Check element tag
  if (element.tagName.toLowerCase() === 'figcaption') {
    return true;
  }

  return false;
}

/**
 * Find caption element near an image or table
 *
 * @param element - Image or table element
 * @param searchParent - Whether to search in parent element
 * @returns Caption element or null
 */
export function findNearbyCaption(
  element: HTMLElement,
  searchParent: boolean = true
): HTMLElement | null {
  // Check next sibling
  const nextSibling = element.nextElementSibling as HTMLElement | null;
  if (nextSibling) {
    if (hasCaptionStyle(nextSibling)) {
      return nextSibling;
    }
    // Check if next sibling text looks like a caption
    const text = nextSibling.textContent?.trim() || '';
    if (looksLikeCaption(text)) {
      return nextSibling;
    }
  }

  // Check previous sibling (captions sometimes come before)
  const prevSibling = element.previousElementSibling as HTMLElement | null;
  if (prevSibling) {
    if (hasCaptionStyle(prevSibling)) {
      return prevSibling;
    }
    const text = prevSibling.textContent?.trim() || '';
    if (looksLikeCaption(text)) {
      return prevSibling;
    }
  }

  // Check parent for caption child
  if (searchParent && element.parentElement) {
    const parent = element.parentElement;

    // Look for figcaption
    const figcaption = parent.querySelector('figcaption');
    if (figcaption) {
      return figcaption as HTMLElement;
    }

    // Look for element with caption class
    const captionElement = parent.querySelector(
      '.caption, .MsoCaption, [class*="caption"]'
    );
    if (captionElement && captionElement !== element) {
      return captionElement as HTMLElement;
    }
  }

  return null;
}

/**
 * Parse all captions from a document and return a map of element IDs to captions
 *
 * @param document - HTML document or element
 * @returns Map of element IDs to parsed captions
 */
export function parseAllCaptions(
  document: Document | HTMLElement
): Map<string, ParsedCaption> {
  const captions = new Map<string, ParsedCaption>();

  // Find all caption elements
  const captionElements = document.querySelectorAll(
    'figcaption, .caption, .MsoCaption, [class*="Caption"]'
  );

  captionElements.forEach((element, index) => {
    const text = element.textContent?.trim() || '';
    const parsed = parseCaption(text);

    if (parsed) {
      // Try to find associated element
      const parent = element.parentElement;
      let associatedId: string | undefined;

      if (parent) {
        const img = parent.querySelector('img');
        const table = parent.querySelector('table');

        if (img) {
          associatedId = img.id || `img-${index}`;
          if (!img.id) img.id = associatedId;
        } else if (table) {
          associatedId = table.id || `table-${index}`;
          if (!table.id) table.id = associatedId;
        }
      }

      if (associatedId) {
        captions.set(associatedId, parsed);
      }
    }
  });

  return captions;
}
