/**
 * RST Word Add-in - Figure Directive Generator
 * Generates RST `.. figure::` directives for images with captions
 *
 * @see https://docutils.sourceforge.io/docs/ref/rst/directives.html#figure
 */

import { FigureOptions } from '../types';

/**
 * Default indentation for directive options and content
 */
const INDENT = '   ';

/**
 * Generate an RST figure directive
 *
 * A figure is an image with a caption and optional legend.
 *
 * @param options - Figure directive options
 * @returns RST figure directive string
 *
 * @example
 * ```typescript
 * const rst = generateFigureDirective({
 *   uri: 'images/architecture.png',
 *   alt: 'System architecture',
 *   width: '600px',
 *   align: 'center',
 *   caption: 'Figure 1: System architecture overview',
 *   legend: 'This diagram shows the main components.'
 * });
 * ```
 */
export function generateFigureDirective(options: FigureOptions): string {
  const lines: string[] = [];

  // Directive declaration
  lines.push(`.. figure:: ${options.uri}`);

  // Image options (inherited from image directive)
  if (options.alt) {
    lines.push(`${INDENT}:alt: ${options.alt}`);
  }

  if (options.height) {
    lines.push(`${INDENT}:height: ${options.height}`);
  }

  if (options.width) {
    lines.push(`${INDENT}:width: ${options.width}`);
  }

  if (options.scale) {
    lines.push(`${INDENT}:scale: ${options.scale}%`);
  }

  // Figure-specific alignment (horizontal only: left, center, right)
  if (options.align && isFigureAlignment(options.align)) {
    lines.push(`${INDENT}:align: ${options.align}`);
  }

  if (options.target) {
    lines.push(`${INDENT}:target: ${options.target}`);
  }

  // Figure-specific options
  if (options.figwidth) {
    lines.push(`${INDENT}:figwidth: ${options.figwidth}`);
  }

  if (options.figclass) {
    lines.push(`${INDENT}:figclass: ${options.figclass}`);
  }

  if (options.figname) {
    lines.push(`${INDENT}:name: ${options.figname}`);
  } else if (options.name) {
    lines.push(`${INDENT}:name: ${options.name}`);
  }

  if (options.class) {
    lines.push(`${INDENT}:class: ${options.class}`);
  }

  // Caption (blank line, then indented paragraph)
  if (options.caption) {
    lines.push('');
    lines.push(`${INDENT}${options.caption}`);
  }

  // Legend (blank line after caption, then indented content)
  if (options.legend) {
    lines.push('');
    // Indent each line of the legend
    const legendLines = options.legend.split('\n');
    for (const line of legendLines) {
      lines.push(`${INDENT}${line}`);
    }
  }

  return lines.join('\n');
}

/**
 * Check if alignment is valid for figures (horizontal only)
 *
 * @param align - Alignment value
 * @returns True if valid figure alignment
 */
export function isFigureAlignment(align: string): boolean {
  return ['left', 'center', 'right'].includes(align.toLowerCase());
}

/**
 * Determine if an image should be a figure based on available data
 *
 * An image becomes a figure when it has:
 * - A caption
 * - A figure number/label
 * - Associated descriptive text
 *
 * @param options - Image/figure options
 * @returns True if should be rendered as figure
 */
export function shouldBeFigure(options: Partial<FigureOptions>): boolean {
  return !!(
    options.caption ||
    options.figureNumber ||
    options.legend ||
    options.figwidth ||
    options.figclass ||
    options.figname
  );
}

/**
 * Convert image options to figure options
 *
 * @param imageOptions - Base image options
 * @param caption - Figure caption
 * @param legend - Optional legend text
 * @returns Figure options
 */
export function imageToFigureOptions(
  imageOptions: {
    uri: string;
    alt?: string;
    width?: string;
    height?: string;
    scale?: string;
    align?: string;
    target?: string;
    class?: string;
    name?: string;
  },
  caption: string,
  legend?: string
): FigureOptions {
  const figureOptions: FigureOptions = {
    uri: imageOptions.uri,
    caption,
  };

  // Copy over compatible options
  if (imageOptions.alt) figureOptions.alt = imageOptions.alt;
  if (imageOptions.width) figureOptions.width = imageOptions.width;
  if (imageOptions.height) figureOptions.height = imageOptions.height;
  if (imageOptions.scale) figureOptions.scale = imageOptions.scale;
  if (imageOptions.target) figureOptions.target = imageOptions.target;
  if (imageOptions.class) figureOptions.class = imageOptions.class;

  // Convert alignment (figure only supports horizontal alignment)
  if (imageOptions.align && isFigureAlignment(imageOptions.align)) {
    figureOptions.align = imageOptions.align as FigureOptions['align'];
  }

  // Use image name as figure name if provided
  if (imageOptions.name) {
    figureOptions.figname = imageOptions.name;
  }

  if (legend) {
    figureOptions.legend = legend;
  }

  return figureOptions;
}

/**
 * Extract figure number from caption text
 *
 * @param caption - Caption text (e.g., "Figure 1: Description")
 * @returns Extracted figure number or undefined
 */
export function extractFigureNumber(caption: string): string | undefined {
  // Match patterns like "Figure 1", "Fig. 2", "Figure 1.2", etc.
  const patterns = [
    /^(?:Figure|Fig\.?)\s+(\d+(?:\.\d+)*)/i,
    /^(\d+(?:\.\d+)*)\s*[:.]\s*/,
  ];

  for (const pattern of patterns) {
    const match = caption.match(pattern);
    if (match) {
      return match[1];
    }
  }

  return undefined;
}

/**
 * Generate a reference name from figure caption
 *
 * @param caption - Figure caption
 * @param figureNumber - Optional figure number
 * @returns Reference name suitable for RST
 */
export function generateFigureRefName(
  caption: string,
  figureNumber?: string
): string {
  if (figureNumber) {
    return `fig-${figureNumber.replace(/\./g, '-')}`;
  }

  // Generate from caption text
  const text = caption
    // Remove figure prefix
    .replace(/^(?:Figure|Fig\.?)\s*\d*[:.]\s*/i, '')
    // Convert to lowercase
    .toLowerCase()
    // Replace non-alphanumeric with hyphens
    .replace(/[^a-z0-9]+/g, '-')
    // Remove leading/trailing hyphens
    .replace(/^-+|-+$/g, '')
    // Limit length
    .substring(0, 50);

  return `fig-${text || 'unnamed'}`;
}

/**
 * Format caption text, optionally removing the figure number prefix
 *
 * @param caption - Original caption
 * @param keepNumber - Whether to keep the figure number
 * @returns Formatted caption
 */
export function formatCaption(caption: string, keepNumber: boolean = true): string {
  if (keepNumber) {
    return caption.trim();
  }

  // Remove "Figure X:" or "Fig. X:" prefix
  return caption
    .replace(/^(?:Figure|Fig\.?)\s*\d+(?:\.\d+)*\s*[:.]\s*/i, '')
    .trim();
}

/**
 * Parse figure information from surrounding HTML context
 *
 * @param imageElement - The image element
 * @param containerElement - The container element (may have caption)
 * @returns Parsed figure options or null if not a figure
 */
export function parseFigureFromHtml(
  imageElement: HTMLImageElement,
  containerElement?: HTMLElement
): FigureOptions | null {
  // Look for caption in various places
  let caption: string | undefined;
  let legend: string | undefined;

  if (containerElement) {
    // Check for figcaption element
    const figcaption = containerElement.querySelector('figcaption');
    if (figcaption) {
      caption = figcaption.textContent?.trim();
    }

    // Check for caption paragraph (Word often uses this)
    const captionPara = containerElement.querySelector('.caption, [class*="Caption"]');
    if (captionPara && !caption) {
      caption = captionPara.textContent?.trim();
    }

    // Check for text after the image that might be a caption
    if (!caption) {
      const nextSibling = imageElement.nextElementSibling;
      if (nextSibling && nextSibling.tagName === 'P') {
        const text = nextSibling.textContent?.trim() || '';
        // Check if it looks like a caption
        if (/^(?:Figure|Fig\.?)\s*\d/i.test(text)) {
          caption = text;
        }
      }
    }
  }

  // If no caption found, not a figure
  if (!caption) {
    return null;
  }

  // Build figure options
  const options: FigureOptions = {
    uri: imageElement.src || imageElement.getAttribute('src') || '',
    caption,
  };

  // Copy image attributes
  const alt = imageElement.alt || imageElement.getAttribute('alt');
  if (alt) options.alt = alt;

  const width = imageElement.width || imageElement.getAttribute('width');
  if (width) options.width = String(width);

  const height = imageElement.height || imageElement.getAttribute('height');
  if (height) options.height = String(height);

  // Extract figure number
  options.figureNumber = extractFigureNumber(caption);

  // Generate reference name
  options.figname = generateFigureRefName(caption, options.figureNumber);

  if (legend) {
    options.legend = legend;
  }

  return options;
}
