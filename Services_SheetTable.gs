/**
 * Services_SheetTable.gs
 *
 * Represents the campaign table stored in a Google Sheet.
 *
 * Why this abstraction exists:
 * - Apps Script ranges are expensive; reading the table once is faster and more predictable.
 * - Header names are user-editable, so we normalize and index them for resilient lookups.
 * - The rest of the codebase should never care about column numbers directly.
 *
 * Conventions:
 * - Header matching is case-insensitive and whitespace-trimmed (via Utils.normalize()).
 * - The first occurrence of a header wins; duplicate detection is handled separately.
 */
class SheetTable {
  /**
   * @param {Object} params
   * @param {GoogleAppsScript.Spreadsheet.Sheet} params.sheet
   * @param {number} params.headerRow
   * @param {string[]} params.headers
   * @param {Object.<string, number>} params.headerIndex
   * @param {{rowNumber:number, values:any[]}[]} params.rows
   * @param {number} params.lastCol
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
   * Loads the header row and all data rows into memory.
   *
   * Performance note:
   * - This method minimizes Range calls (1 for header + 1 for data block).
   * - Intended to be called once per orchestrator run.
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

    /**
     * Normalized header → zero-based column index.
     * First occurrence wins so that duplicates can be reported but do not create ambiguity here.
     */
    const headerIndex = {};
    headers.forEach((h, i) => {
      const key = Utils.normalize(h);
      if (!key) return;
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
   * Finds duplicate header labels (case-insensitive).
   *
   * Why it matters:
   * - Duplicate headers make the sheet ambiguous and can route data to the wrong field.
   * - The orchestrator uses this to block sending until the sheet is corrected.
   *
   * @returns {string[]}
   *   Human-readable problem descriptions.
   */
  validateNoDuplicateHeaders() {
    const seen = new Map();

    this.headers.forEach((h, i) => {
      const key = Utils.normalize(h);
      if (!key) return;

      if (!seen.has(key)) seen.set(key, []);
      seen.get(key).push(i + 1); // store 1-based column numbers for user readability
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
   * Returns the zero-based column index for a header name (or null if missing).
   *
   * This method intentionally supports a few common variations so that the system
   * remains tolerant to minor header spelling differences.
   *
   * @param {string} name
   * @returns {number|null}
   */
  getIndex(name) {
    const key = Utils.normalize(name);

    /**
     * Header normalization supports variations without forcing users to rename columns.
     * If you add a new canonical header, also consider listing its common variants here.
     */
    if (key === 'replyto' || key === 'reply to' || key === 'reply_to') {
      return this.headerIndex['replyto'] ??
        this.headerIndex['reply to'] ??
        this.headerIndex['reply_to'] ??
        null;
    }

    if (key === 'noreply' || key === 'no reply' || key === 'no_reply') {
      return this.headerIndex['noreply'] ??
        this.headerIndex['no reply'] ??
        this.headerIndex['no_reply'] ??
        null;
    }

    if (key === 'fromname' || key === 'from name') {
      return this.headerIndex['fromname'] ??
        this.headerIndex['from name'] ??
        null;
    }

    if (key === 'fromemail' || key === 'from email') {
      return this.headerIndex['fromemail'] ??
        this.headerIndex['from email'] ??
        null;
    }

    return this.headerIndex[key] ?? null;
  }

  /**
   * Reads a single row by absolute row number.
   *
   * Use case:
   * - The test row can live outside the data block (e.g., above the header),
   *   so it cannot be addressed through this.rows.
   *
   * @param {number} rowNumber
   * @returns {{rowNumber:number, values:any[]}}
   */
  readAbsoluteRow(rowNumber) {
    const values = this.sheet.getRange(rowNumber, 1, 1, this.lastCol).getValues()[0];
    return { rowNumber, values };
  }
}
