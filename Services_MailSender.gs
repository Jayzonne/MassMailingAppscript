/**
 * Services_MailSender.gs
 *
 * MailSender encapsulates operational concerns during sending:
 * - Marking rows as sent (immediately after a successful send)
 * - Optionally clearing the "To send" checkbox
 * - Optionally writing a "sentAt" timestamp
 * - Applying throttling delays between sends
 *
 * This class does NOT:
 * - build email options (EmailComposer does that)
 * - render templates (TemplateRenderer does that)
 * - decide which rows to send (Orchestrator does that)
 *
 * It focuses only on side effects and execution pacing.
 */
class MailSender {

  /**
   * Creates a new MailSender.
   *
   * @param {Object} params
   * @param {GoogleAppsScript.Spreadsheet.Sheet} params.sheet
   *   Active sheet used to update row state (Sent / To send / sentAt).
   * @param {Object} params.config
   *   Application configuration (APP_CONFIG).
   */
  constructor({ sheet, config }) {
    this.sheet = sheet;
    this.config = config;
  }

  /**
   * Marks a row as successfully sent and updates related state immediately.
   *
   * Updates performed (depending on config):
   * - Set "Sent" checkbox to TRUE
   * - Optionally clear "To send" checkbox
   * - Optionally write the current timestamp into "sentAt"
   *
   * SpreadsheetApp.flush() is called to force UI/state updates promptly.
   * This is important when long batches are running, so users can observe progress.
   *
   * @param {SheetTable} table
   *   Parsed sheet table for header index resolution.
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

    if (this.config.marking.writeSentTimestamp && idxSentAt != null) {
      this.sheet.getRange(rowNumber, idxSentAt + 1).setValue(new Date());
    }

    SpreadsheetApp.flush();
  }

  /**
   * Applies throttling between send attempts.
   *
   * A random delay is used to:
   * - reduce Gmail anti-spam triggers
   * - prevent rate-limit errors
   * - avoid burst-like traffic patterns
   *
   * The delay bounds are defined in config.throttling.
   */
  throttle() {
    Utils.sleepRandomSeconds(
      this.config.throttling.secondsMin,
      this.config.throttling.secondsMax
    );
  }
}
