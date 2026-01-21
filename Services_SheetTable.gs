/**
 * Services_SheetTable.gs
 *
 * SheetTable is a structured in-memory representation of a Google Sheet table.
 *
 * Responsibilities:
 * - Read the header row and build a normalized header index
 * - Read all data rows below the header row into a consistent format
 * - Provide safe header lookup (with synonyms) for downstream services
 * - Validate structural constraints (e.g., no duplicate headers)
 *
 * This class is intentionally focused on "table structure" only.
 * It does not contain business logic such as email sending or templating.
 */
class SheetTable {

  /**
   * Creates a new SheetTable instance.
   *
   * @param {Object} params
   * @param {GoogleAppsScript.Spreadsheet.Sheet} params.sheet
   *   Source sheet used as the table backend.
   * @param {number} params.headerRow
   *   1-based row index where headers are located.
   * @param {string[]} params.headers
   *   Raw header values as displayed in the sheet (trimmed).
   * @param {Object.<string, number>} params.headerIndex
   *   Map of normalized header name -> 0-based column index.
   * @param {Array<{rowNumber:number, values:any[]}>} params.rows
   *   Data rows (only rows strictly below the header row).
   * @param {number} params.lastCol
   *   Last column count at read time (used for absolute row reads).
   */
  constructor({ sheet, headerRow, headers, headerIndex, rows, lastCol }) {
    this.sheet = sheet;
    this.headerRow = headerRow;
    this.headers = headers;
    this.headerIndex = headerIndex;
    this.rows = rows;
    this.lastCol = lastCol;
  }

  /**
   * Factory method to build a SheetTable from a Google Sheet.
   *
   * Reads:
   * - header row cells across all columns up to `sheet.getLastColumn()`
   * - every data row below the header row up to `sheet.getLastRow()`
   *
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   * @param {number} headerRow
   * @returns {SheetTable}
   */
  static fromSheet(sheet, headerRow) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    const headerValues = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
    const headers = headerValues.map(h => String(h || '').trim());

    const headerIndex = {};
    headers.forEach((h, i) => {
      const key = Utils.normalize(h);
      if (!key) return;

      // Keep first occurrence; duplicates are detected by validateNoDuplicateHeaders()
      if (headerIndex[key] == null) headerIndex[key] = i;
    });

    const numDataRows = Math.max(0, lastRow - headerRow);
    const dataValues = numDataRows
      ? sheet.getRange(headerRow + 1, 1, numDataRows, lastCol).getValues()
      : [];

    const rows = dataValues.map((vals, i) => ({
      rowNumber: headerRow + 1 + i,
      values: vals,
    }));

    return new SheetTable({ sheet, headerRow, headers, headerIndex, rows, lastCol });
  }

  /**
   * Validates that the table does not contain duplicate headers
   * (case-insensitive and whitespace-insensitive).
   *
   * Duplicate headers create ambiguous behavior and must be blocked.
   *
   * @returns {string[]} List of human-readable validation errors (empty if ok).
   */
  validateNoDuplicateHeaders() {
    const seen = new Map();

    this.headers.forEach((h, i) => {
      const key = Utils.normalize(h);
      if (!key) return;

      if (!seen.has(key)) seen.set(key, []);
      // store 1-based column numbers for user-facing messages
      seen.get(key).push(i + 1);
    });

    const problems = [];
    for (const [key, cols] of seen.entries()) {
      if (cols.length > 1) {
        problems.push(`Duplicate header "${key}" found in columns: ${cols.join(', ')}`);
      }
    }

    return problems;
  }

  /**
   * Returns the 0-based column index for a given header name.
   *
   * - Matching is normalized (lowercase + trim)
   * - Supports common synonyms used in the template:
   *   - replyTo / reply to / reply_to
   *   - noReply / no reply / no_reply
   *   - fromName / from name
   *   - fromEmail / from email
   *
   * @param {string} name
   *   Header name to look up.
   * @returns {number|null}
   *   0-based column index, or null if header not found.
   */
  getIndex(name) {
    const key = Utils.normalize(name);

    // Synonyms mapping to support flexible header naming conventions
    if (key === 'replyto' || key === 'reply to' || key === 'reply_to') {
      return this.headerIndex['replyto'] ?? this.headerIndex['reply to'] ?? this.headerIndex['reply_to'] ?? null;
    }
    if (key === 'noreply' || key === 'no reply' || key === 'no_reply') {
      return this.headerIndex['noreply'] ?? this.headerIndex['no reply'] ?? this.headerIndex['no_reply'] ?? null;
    }
    if (key === 'fromname' || key === 'from name') {
      return this.headerIndex['fromname'] ?? this.headerIndex['from name'] ?? null;
    }
    if (key === 'fromemail' || key === 'from email') {
      return this.headerIndex['fromemail'] ?? this.headerIndex['from email'] ?? null;
    }

    return this.headerIndex[key] ?? null;
  }

  /**
   * Reads a single absolute row from the sheet across the table width.
   *
   * This is primarily used for the "Test email" workflow where the test row
   * can live above the header row, and therefore is not part of `this.rows`.
   *
   * @param {number} rowNumber
   *   1-based sheet row number to read.
   * @returns {{rowNumber:number, values:any[]}}
   */
  readAbsoluteRow(rowNumber) {
    const values = this.sheet.getRange(rowNumber, 1, 1, this.lastCol).getValues()[0];
    return { rowNumber, values };
  }
}
