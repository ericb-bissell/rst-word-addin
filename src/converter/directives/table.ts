/**
 * RST Word Add-in - Table Directive Generator
 * Generates RST `.. table::` directives and grid/simple tables
 *
 * @see https://docutils.sourceforge.io/docs/ref/rst/directives.html#table
 */

import { TableOptions, TableData, TableRow, TableCell } from '../types';

/**
 * Default indentation for directive options
 */
const INDENT = '   ';

/**
 * Generate an RST table directive with grid table content
 *
 * @param data - Table data including rows and options
 * @returns RST table directive string
 *
 * @example
 * ```typescript
 * const rst = generateTableDirective({
 *   rows: [
 *     { cells: [{ content: 'Name' }, { content: 'Age' }], isHeader: true },
 *     { cells: [{ content: 'Alice' }, { content: '30' }] },
 *   ],
 *   options: { caption: 'Table 1: User Data', align: 'center' }
 * });
 * ```
 */
export function generateTableDirective(data: TableData): string {
  const lines: string[] = [];
  const { rows, options } = data;

  // If table has caption or options, use directive wrapper
  const useDirective = !!(
    options.caption ||
    options.align ||
    options.width ||
    options.widths ||
    options.class ||
    options.name
  );

  if (useDirective) {
    // Directive declaration with optional caption
    if (options.caption) {
      lines.push(`.. table:: ${options.caption}`);
    } else {
      lines.push('.. table::');
    }

    // Add options
    if (options.align) {
      lines.push(`${INDENT}:align: ${options.align}`);
    }

    if (options.width) {
      lines.push(`${INDENT}:width: ${options.width}`);
    }

    if (options.widths) {
      const widthsStr = Array.isArray(options.widths)
        ? options.widths.join(' ')
        : options.widths;
      lines.push(`${INDENT}:widths: ${widthsStr}`);
    }

    if (options.class) {
      lines.push(`${INDENT}:class: ${options.class}`);
    }

    if (options.name) {
      lines.push(`${INDENT}:name: ${options.name}`);
    }

    lines.push('');

    // Generate grid table with indentation
    const gridTable = generateGridTable(rows, options.hasHeader);
    const indentedTable = gridTable
      .split('\n')
      .map((line) => INDENT + line)
      .join('\n');
    lines.push(indentedTable);
  } else {
    // Just generate the grid table without directive wrapper
    lines.push(generateGridTable(rows, options.hasHeader));
  }

  return lines.join('\n');
}

/**
 * Generate an RST grid table
 *
 * Grid tables use +, -, |, and = characters to draw cell borders.
 *
 * @param rows - Table rows
 * @param hasHeader - Whether first row(s) are headers
 * @returns Grid table string
 */
export function generateGridTable(rows: TableRow[], hasHeader?: boolean): string {
  if (rows.length === 0) {
    return '';
  }

  // Calculate column widths
  const columnWidths = calculateColumnWidths(rows);
  const numColumns = columnWidths.length;

  const lines: string[] = [];

  // Generate separator line
  const normalSeparator = generateSeparator(columnWidths, '-');
  const headerSeparator = generateSeparator(columnWidths, '=');

  // Top border
  lines.push(normalSeparator);

  // Process rows
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    const isHeaderRow = row.isHeader || (hasHeader && rowIndex === 0);

    // Generate row content (may span multiple lines for wrapped text)
    const rowLines = generateRowLines(row, columnWidths, numColumns);
    lines.push(...rowLines);

    // Add separator after row
    if (isHeaderRow) {
      lines.push(headerSeparator);
    } else {
      lines.push(normalSeparator);
    }
  }

  return lines.join('\n');
}

/**
 * Generate a simple RST table (easier to read but less flexible)
 *
 * @param rows - Table rows
 * @param hasHeader - Whether first row is a header
 * @returns Simple table string
 */
export function generateSimpleTable(rows: TableRow[], hasHeader?: boolean): string {
  if (rows.length === 0) {
    return '';
  }

  const columnWidths = calculateColumnWidths(rows);
  const lines: string[] = [];

  // Generate separator
  const separator = columnWidths.map((w) => '='.repeat(w)).join('  ');

  // Top border
  lines.push(separator);

  // Process rows
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    const isHeaderRow = row.isHeader || (hasHeader && rowIndex === 0);

    // Generate cell content
    const cellContents = row.cells.map((cell, colIndex) => {
      const width = columnWidths[colIndex] || 10;
      return padCell(cell.content, width, cell.align);
    });

    lines.push(cellContents.join('  '));

    // Add header separator after header row
    if (isHeaderRow) {
      lines.push(separator);
    }
  }

  // Bottom border
  lines.push(separator);

  return lines.join('\n');
}

/**
 * Calculate column widths based on content
 *
 * @param rows - Table rows
 * @param minWidth - Minimum column width
 * @param maxWidth - Maximum column width
 * @returns Array of column widths
 */
export function calculateColumnWidths(
  rows: TableRow[],
  minWidth: number = 3,
  maxWidth: number = 40
): number[] {
  if (rows.length === 0) {
    return [];
  }

  // Find maximum number of columns
  const numColumns = Math.max(...rows.map((row) => row.cells.length));

  // Initialize widths with minimum
  const widths: number[] = new Array(numColumns).fill(minWidth);

  // Calculate max width for each column
  for (const row of rows) {
    for (let colIndex = 0; colIndex < row.cells.length; colIndex++) {
      const cell = row.cells[colIndex];
      const contentWidth = getContentWidth(cell.content);
      widths[colIndex] = Math.max(widths[colIndex], contentWidth);
    }
  }

  // Apply max width limit
  return widths.map((w) => Math.min(w, maxWidth));
}

/**
 * Get the display width of content (considering line breaks)
 *
 * @param content - Cell content
 * @returns Maximum line width
 */
function getContentWidth(content: string): number {
  const lines = content.split('\n');
  return Math.max(...lines.map((line) => line.length), 1);
}

/**
 * Generate a table separator line
 *
 * @param columnWidths - Array of column widths
 * @param char - Character to use for horizontal lines ('-' or '=')
 * @returns Separator line
 */
function generateSeparator(columnWidths: number[], char: string): string {
  const segments = columnWidths.map((width) => char.repeat(width + 2));
  return '+' + segments.join('+') + '+';
}

/**
 * Generate lines for a single row (handles text wrapping)
 *
 * @param row - Table row
 * @param columnWidths - Column widths
 * @param numColumns - Total number of columns
 * @returns Array of row content lines
 */
function generateRowLines(
  row: TableRow,
  columnWidths: number[],
  numColumns: number
): string[] {
  // Split each cell content into lines
  const cellLines: string[][] = [];

  for (let colIndex = 0; colIndex < numColumns; colIndex++) {
    const cell = row.cells[colIndex];
    const width = columnWidths[colIndex] || 10;

    if (cell) {
      // Wrap content to fit column width
      const wrapped = wrapText(cell.content, width);
      cellLines.push(wrapped);
    } else {
      cellLines.push(['']);
    }
  }

  // Find maximum number of lines in any cell
  const maxLines = Math.max(...cellLines.map((lines) => lines.length), 1);

  // Generate output lines
  const outputLines: string[] = [];

  for (let lineIndex = 0; lineIndex < maxLines; lineIndex++) {
    const segments: string[] = [];

    for (let colIndex = 0; colIndex < numColumns; colIndex++) {
      const width = columnWidths[colIndex] || 10;
      const cell = row.cells[colIndex];
      const lines = cellLines[colIndex] || [''];
      const lineContent = lines[lineIndex] || '';
      const align = cell?.align || 'left';

      segments.push(' ' + padCell(lineContent, width, align) + ' ');
    }

    outputLines.push('|' + segments.join('|') + '|');
  }

  return outputLines;
}

/**
 * Wrap text to fit within a given width
 *
 * @param text - Text to wrap
 * @param width - Maximum width
 * @returns Array of wrapped lines
 */
function wrapText(text: string, width: number): string[] {
  if (!text) {
    return [''];
  }

  // Handle existing line breaks
  const paragraphs = text.split('\n');
  const result: string[] = [];

  for (const paragraph of paragraphs) {
    if (paragraph.length <= width) {
      result.push(paragraph);
    } else {
      // Word wrap
      const words = paragraph.split(' ');
      let currentLine = '';

      for (const word of words) {
        if (currentLine.length === 0) {
          currentLine = word;
        } else if (currentLine.length + 1 + word.length <= width) {
          currentLine += ' ' + word;
        } else {
          result.push(currentLine);
          currentLine = word;
        }
      }

      if (currentLine) {
        result.push(currentLine);
      }
    }
  }

  return result.length > 0 ? result : [''];
}

/**
 * Pad cell content to width with alignment
 *
 * @param content - Cell content
 * @param width - Target width
 * @param align - Alignment (left, center, right)
 * @returns Padded string
 */
function padCell(
  content: string,
  width: number,
  align?: 'left' | 'center' | 'right'
): string {
  const text = content || '';

  if (text.length >= width) {
    return text.substring(0, width);
  }

  const padding = width - text.length;

  switch (align) {
    case 'right':
      return ' '.repeat(padding) + text;
    case 'center':
      const leftPad = Math.floor(padding / 2);
      const rightPad = padding - leftPad;
      return ' '.repeat(leftPad) + text + ' '.repeat(rightPad);
    case 'left':
    default:
      return text + ' '.repeat(padding);
  }
}

/**
 * Parse table from HTML table element
 *
 * @param tableElement - HTML table element
 * @returns Parsed table data
 */
export function parseHtmlTable(tableElement: HTMLTableElement): TableData {
  const rows: TableRow[] = [];
  const options: TableOptions = {};

  // Check for caption
  const caption = tableElement.querySelector('caption');
  if (caption) {
    options.caption = caption.textContent?.trim();

    // Try to extract table number
    const numberMatch = options.caption?.match(/^(?:Table|Tbl\.?)\s+(\d+(?:\.\d+)*)/i);
    if (numberMatch) {
      options.tableNumber = numberMatch[1];
    }
  }

  // Parse thead
  const thead = tableElement.querySelector('thead');
  if (thead) {
    const headerRows = thead.querySelectorAll('tr');
    headerRows.forEach((tr) => {
      rows.push(parseTableRow(tr, true));
    });
    options.hasHeader = true;
  }

  // Parse tbody
  const tbody = tableElement.querySelector('tbody');
  const bodyElement = tbody || tableElement;
  const bodyRows = bodyElement.querySelectorAll(':scope > tr');

  bodyRows.forEach((tr, index) => {
    // If no thead, first row might be header (check for th elements)
    const hasThElements = tr.querySelectorAll('th').length > 0;
    const isHeader = !thead && index === 0 && hasThElements;

    if (isHeader) {
      options.hasHeader = true;
    }

    rows.push(parseTableRow(tr, isHeader));
  });

  // Parse table attributes
  const align = tableElement.getAttribute('align');
  if (align && ['left', 'center', 'right'].includes(align)) {
    options.align = align as TableOptions['align'];
  }

  const width = tableElement.getAttribute('width');
  if (width) {
    options.width = width;
  }

  return { rows, options };
}

/**
 * Parse a table row from HTML
 *
 * @param tr - HTML table row element
 * @param isHeader - Whether this is a header row
 * @returns Parsed table row
 */
function parseTableRow(tr: Element, isHeader: boolean): TableRow {
  const cells: TableCell[] = [];

  const cellElements = tr.querySelectorAll('td, th');
  cellElements.forEach((cell) => {
    const tableCell: TableCell = {
      content: cell.textContent?.trim() || '',
    };

    // Check for colspan
    const colspan = cell.getAttribute('colspan');
    if (colspan && parseInt(colspan) > 1) {
      tableCell.colspan = parseInt(colspan);
    }

    // Check for rowspan
    const rowspan = cell.getAttribute('rowspan');
    if (rowspan && parseInt(rowspan) > 1) {
      tableCell.rowspan = parseInt(rowspan);
    }

    // Check for alignment
    const align = cell.getAttribute('align');
    if (align && ['left', 'center', 'right'].includes(align)) {
      tableCell.align = align as TableCell['align'];
    }

    cells.push(tableCell);
  });

  return { cells, isHeader };
}

/**
 * Generate a reference name from table caption
 *
 * @param caption - Table caption
 * @param tableNumber - Optional table number
 * @returns Reference name suitable for RST
 */
export function generateTableRefName(
  caption: string,
  tableNumber?: string
): string {
  if (tableNumber) {
    return `tbl-${tableNumber.replace(/\./g, '-')}`;
  }

  // Generate from caption text
  const text = caption
    .replace(/^(?:Table|Tbl\.?)\s*\d*[:.]\s*/i, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .substring(0, 50);

  return `tbl-${text || 'unnamed'}`;
}
