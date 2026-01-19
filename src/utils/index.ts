/**
 * RST Word Add-in - Utilities Module
 * Exports all utility functions
 */

// Caption parser
export {
  parseCaption,
  looksLikeCaption,
  detectCaptionType,
  extractCaptionNumber,
  extractCaptionText,
  formatCaptionForRst,
  generateRefLabel,
  hasCaptionStyle,
  findNearbyCaption,
  parseAllCaptions,
} from './caption-parser';
