/**
 * Services_MailSender.gs
 *
 * Applies side effects on the sheet during mail operations:
 * - Marks rows as sent (checkbox + timestamp)
 * - Clears the "To send" flag after success (optional)
 * - Flushes updates so progress is visible while long batches are running
 * - Throttles between sends to reduce quota / anti-spam / burst behavior issues
 *
 * This service does not send email. It only updates sheet state and pacing.
 */
class MailSender {
  /**
   * @param {Object} params
   * @param {GoogleAppsScript.Spreadsheet.Sheet} params.sheet
   *   The sheet where status columns live.
   * @param {Object} params.config
   *   APP_CONFIG (used for header names and behavior toggles).
   */
  constructor({ sheet, config }) {
    this.sheet = sheet;
    this.config = config;
  }

  /**
   * Persists a "successful send" state on a specific row.
   *
   * Why this is immediate:
   * - If a batch is long and partially fails, users still get accurate progress feedback.
   * - If execution stops mid-run (quota/time), already-sent rows remain correctly marked.
   *
   * @param {SheetTable} table
   *   Parsed table used to resolve status column indices.
   * @param {number} rowNumber
   *   Absolute 1-based row number to update.
   */
  markSentNow(table, rowNumber) {
    const idxSent = table.getIndex(this.config.headers.sent);
    const idxToSend = table.getIndex(this.config.headers.toSend);
    const idxSentAt = table.getIndex(this.config.headers.sentAt);

    if (idxSent != null) {
      this.sheet.getRange(rowNumber, idxSent + 1).setValue(true);
    }

    if (this.config.marking.clearToSendAfterSent && idxToSend != null) {
      this.sheet.getRange(rowNumber, idxToSend + 1).setValue(false);
    }

    /**
     * Timestamp is formatted explicitly to avoid locale-dependent rendering.
     * The format should match what the template builder applies in reconstruction.
     */
    if (this.config.marking.writeSentTimestamp && idxSentAt != null) {
      const now = new Date();
      const cell = this.sheet.getRange(rowNumber, idxSentAt + 1);
      cell.setValue(now);
      cell.setNumberFormat(this.config.sentAtNumberFormat || 'yyyy/MM/dd - HH:mm:ss');
    }

    SpreadsheetApp.flush();
  }

  /**
   * Adds a random delay between sends.
   *
   * Randomized throttling avoids burst-like traffic patterns and
   * reduces the probability of rate-limit / anti-spam issues.
   */
  throttle() {
    Utils.sleepRandomSeconds(
      this.config.throttling.secondsMin,
      this.config.throttling.secondsMax
    );
  }
}
