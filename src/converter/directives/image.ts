/**
 * RST Word Add-in - Image Directive Generator
 * Generates RST `.. image::` directives
 *
 * @see https://docutils.sourceforge.io/docs/ref/rst/directives.html#image
 */

import { ImageOptions } from '../types';

/**
 * Default indentation for directive options
 */
const INDENT = '   ';

/**
 * Generate an RST image directive
 *
 * @param options - Image directive options
 * @returns RST image directive string
 *
 * @example
 * ```typescript
 * const rst = generateImageDirective({
 *   uri: 'images/photo.png',
 *   alt: 'A photo',
 *   width: '400px',
 *   align: 'center'
 * });
 * // Returns:
 * // .. image:: images/photo.png
 * //    :alt: A photo
 * //    :width: 400px
 * //    :align: center
 * ```
 */
export function generateImageDirective(options: ImageOptions): string {
  const lines: string[] = [];

  // Directive declaration
  lines.push(`.. image:: ${options.uri}`);

  // Add options in a consistent order
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

  if (options.align) {
    lines.push(`${INDENT}:align: ${options.align}`);
  }

  if (options.target) {
    lines.push(`${INDENT}:target: ${options.target}`);
  }

  if (options.class) {
    lines.push(`${INDENT}:class: ${options.class}`);
  }

  if (options.name) {
    lines.push(`${INDENT}:name: ${options.name}`);
  }

  if (options.loading) {
    lines.push(`${INDENT}:loading: ${options.loading}`);
  }

  return lines.join('\n');
}

/**
 * Parse image options from HTML img element attributes
 *
 * @param element - HTML image element or attributes object
 * @param imagePath - Path to the image file
 * @returns Parsed image options
 */
export function parseImageOptions(
  element: HTMLImageElement | Record<string, string>,
  imagePath: string
): ImageOptions {
  const options: ImageOptions = {
    uri: imagePath,
  };

  // Handle both HTMLImageElement and plain object
  const getAttribute = (name: string): string | null => {
    if (element instanceof HTMLImageElement) {
      return element.getAttribute(name);
    }
    return element[name] || null;
  };

  // Alt text
  const alt = getAttribute('alt');
  if (alt) {
    options.alt = alt;
  }

  // Dimensions
  const width = getAttribute('width');
  if (width) {
    options.width = normalizeSize(width);
  }

  const height = getAttribute('height');
  if (height) {
    options.height = normalizeSize(height);
  }

  // Style attribute parsing for additional properties
  const style = getAttribute('style');
  if (style) {
    const styleOptions = parseStyleAttribute(style);

    if (styleOptions.width && !options.width) {
      options.width = styleOptions.width;
    }

    if (styleOptions.height && !options.height) {
      options.height = styleOptions.height;
    }

    if (styleOptions.align) {
      options.align = styleOptions.align as ImageOptions['align'];
    }
  }

  // Data attributes that might contain additional info
  const dataAlign = getAttribute('data-align');
  if (dataAlign && isValidAlignment(dataAlign)) {
    options.align = dataAlign as ImageOptions['align'];
  }

  // Check for hyperlink (if image is wrapped in an anchor)
  const dataTarget = getAttribute('data-target');
  if (dataTarget) {
    options.target = dataTarget;
  }

  return options;
}

/**
 * Normalize size value to include unit if missing
 *
 * @param size - Size value (e.g., "400", "400px", "50%")
 * @returns Normalized size with unit
 */
export function normalizeSize(size: string): string {
  const trimmed = size.trim();

  // Already has a unit
  if (/[a-z%]$/i.test(trimmed)) {
    return trimmed;
  }

  // Numeric only - assume pixels
  if (/^\d+(\.\d+)?$/.test(trimmed)) {
    return `${trimmed}px`;
  }

  return trimmed;
}

/**
 * Parse CSS style attribute for image-related properties
 *
 * @param style - CSS style string
 * @returns Parsed style options
 */
export function parseStyleAttribute(style: string): {
  width?: string;
  height?: string;
  align?: string;
} {
  const result: { width?: string; height?: string; align?: string } = {};

  // Parse style properties
  const properties = style.split(';').map((p) => p.trim()).filter(Boolean);

  for (const prop of properties) {
    const [name, value] = prop.split(':').map((s) => s.trim());

    switch (name.toLowerCase()) {
      case 'width':
        result.width = value;
        break;
      case 'height':
        result.height = value;
        break;
      case 'float':
        if (value === 'left' || value === 'right') {
          result.align = value;
        }
        break;
      case 'margin-left':
      case 'margin-right':
        // Check for auto margins (centering)
        if (value === 'auto') {
          result.align = 'center';
        }
        break;
      case 'text-align':
        if (value === 'left' || value === 'center' || value === 'right') {
          result.align = value;
        }
        break;
      case 'vertical-align':
        if (value === 'top' || value === 'middle' || value === 'bottom') {
          result.align = value;
        }
        break;
    }
  }

  return result;
}

/**
 * Check if a string is a valid RST image alignment value
 *
 * @param value - Alignment value to check
 * @returns True if valid alignment
 */
export function isValidAlignment(value: string): boolean {
  const validAlignments = ['top', 'middle', 'bottom', 'left', 'center', 'right'];
  return validAlignments.includes(value.toLowerCase());
}

/**
 * Generate a unique image filename
 *
 * @param index - Image index in the document
 * @param format - Image format (e.g., "png", "jpg")
 * @param prefix - Filename prefix (default: "image")
 * @returns Generated filename
 */
export function generateImageFilename(
  index: number,
  format: string,
  prefix: string = 'image'
): string {
  const paddedIndex = String(index).padStart(3, '0');
  const normalizedFormat = format.toLowerCase().replace('jpeg', 'jpg');
  return `${prefix}_${paddedIndex}.${normalizedFormat}`;
}

/**
 * Get image format from data URI or file extension
 *
 * @param source - Data URI or filename
 * @returns Image format (e.g., "png", "jpg", "gif")
 */
export function getImageFormat(source: string): string {
  // Check for data URI
  const dataUriMatch = source.match(/^data:image\/(\w+);/i);
  if (dataUriMatch) {
    return dataUriMatch[1].toLowerCase().replace('jpeg', 'jpg');
  }

  // Check file extension
  const extMatch = source.match(/\.(\w+)$/i);
  if (extMatch) {
    return extMatch[1].toLowerCase().replace('jpeg', 'jpg');
  }

  // Default to PNG
  return 'png';
}
