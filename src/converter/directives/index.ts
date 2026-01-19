/**
 * RST Word Add-in - Directives Module
 * Exports all directive generators
 */

// Image directive
export {
  generateImageDirective,
  parseImageOptions,
  normalizeSize,
  parseStyleAttribute,
  isValidAlignment,
  generateImageFilename,
  getImageFormat,
} from './image';

// Figure directive
export {
  generateFigureDirective,
  isFigureAlignment,
  shouldBeFigure,
  imageToFigureOptions,
  extractFigureNumber,
  generateFigureRefName,
  formatCaption,
  parseFigureFromHtml,
} from './figure';

// Table directive
export {
  generateTableDirective,
  generateGridTable,
  generateSimpleTable,
  calculateColumnWidths,
  parseHtmlTable,
  generateTableRefName,
} from './table';

// Contents directive
export {
  generateContentsDirective,
  isTocElement,
  parseTocOptions,
  recommendBacklinks,
  containsTocField,
  extractTocTitle,
} from './contents';

// Custom directives (rst_* styles)
export {
  isRstDirectiveStyle,
  extractDirectiveName,
  parseCustomDirective,
  generateCustomDirective,
  normalizeDirectiveName,
  getKnownDirectives,
  isKnownDirective,
  suggestDirectiveName,
  mergeConsecutiveDirectives,
} from './custom';
