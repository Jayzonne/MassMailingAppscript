/**
 * Services_EmailComposer.gs
 *
 * EmailComposer builds:
 * - Gmail sending parameters (recipient list, subject, Gmail options)
 * - Template variable maps used by the TemplateRenderer
 *
 * Responsibilities:
 * - Extract email-related values from a table row using canonical headers
 * - Normalize recipient lists (supports multiple recipients per cell)
 * - Apply "No Reply" behavior when enabled
 * - Ensure reserved email headers are excluded from template variables
 *
 * This class does not send emails and does not render templates.
 * It only prepares structured inputs for those operations.
 */
class EmailComposer {

  /**
   * Creates a new EmailComposer.
   *
   * @param {Object} config
   *   Application configuration (APP_CONFIG). Used for:
   *   - reserved header names
   *   - no-reply sender configuration
   */
  constructor(config) {
    this.config = config;
  }

  /**
   * Builds Gmail "sendEmail" arguments from a given table row.
   *
   * Data source:
   * - Email configuration is read from reserved (blue) columns only.
   * - Multiple addresses are supported in Email/cc/bcc/replyTo fields.
   *
   * Behavior:
   * - If "No Reply" is enabled, the sender is forced to `config.noReplyFromEmail`
   *   (must be an allowed Gmail alias).
   * - Otherwise, replyTo/fromEmail/fromName are applied when provided.
   * - Per-row subject overrides the default subject if a "subject" column exists.
   *
   * @param {SheetTable} table
   *   Parsed sheet table providing header lookup.
   * @param {{rowNumber:number, values:any[]}} row
   *   Row object containing raw cell values.
   * @param {string} defaultSubject
   *   Fallback subject read from the sheet-level configuration cell.
   * @returns {{to:string, subject:string, options:Object}}
   *   Object ready to be passed to GmailApp.sendEmail(to, subject, body, options).
   */
  buildEmailOptions(table, row, defaultSubject) {
    const idxTo = table.getIndex('email');
    const idxCc = table.getIndex('cc');
    const idxBcc = table.getIndex('bcc');
    const idxReplyTo = table.getIndex('replyto');
    const idxNoReply = table.getIndex('noreply');
    const idxFromName = table.getIndex('fromname');
    const idxFromEmail = table.getIndex('fromemail');
    const idxSubject = table.getIndex('subject');

    const to = Utils.normalizeEmailList(row.values[idxTo]);
    const cc = idxCc != null ? Utils.normalizeEmailList(row.values[idxCc]) : '';
    const bcc = idxBcc != null ? Utils.normalizeEmailList(row.values[idxBcc]) : '';
    const replyTo = idxReplyTo != null ? Utils.normalizeEmailList(row.values[idxReplyTo]) : '';
    const noReply = idxNoReply != null ? Utils.asBool(row.values[idxNoReply]) : false;

    const fromName = idxFromName != null ? String(row.values[idxFromName] || '').trim() : '';
    const fromEmail = idxFromEmail != null ? String(row.values[idxFromEmail] || '').trim() : '';
    const rowSubject = idxSubject != null ? String(row.values[idxSubject] || '').trim() : '';

    const subject = rowSubject || defaultSubject;

    /** @type {GoogleAppsScript.Gmail.GmailAdvancedSendEmailOptions} */
    const options = {};

    if (cc) options.cc = cc;
    if (bcc) options.bcc = bcc;

    if (noReply) {
      /**
       * No-reply behavior:
       * - Force sender address
       * - Avoid setting replyTo
       *
       * Note: Gmail will only accept this if the address is configured
       * as an allowed alias ("Send mail as") in the current account.
       */
      options.from = this.config.noReplyFromEmail;
      options.name = fromName || 'No reply';
    } else {
      if (replyTo) options.replyTo = replyTo;
      if (fromEmail) options.from = fromEmail;
      if (fromName) options.name = fromName;
    }

    return { to, subject, options };
  }

  /**
   * Builds the variable map used to render the Google Docs template.
   *
   * Rule:
   * - Every non-reserved column header becomes a template variable key
   * - Reserved email/config headers (blue columns) are excluded
   * - Control columns (toSend/sent/sentAt) are excluded
   *
   * Example:
   * - Sheet header "Topic1" => template placeholder "$Topic1$"
   *
   * @param {SheetTable} table
   * @param {{rowNumber:number, values:any[]}} row
   * @returns {Object.<string, string>}
   *   Map of template variables.
   */
  buildTemplateVars(table, row) {
    const reserved = new Set(this.config.reservedEmailHeaders.map(Utils.normalize));
    reserved.add(Utils.normalize(this.config.headers.toSend));
    reserved.add(Utils.normalize(this.config.headers.sent));
    reserved.add(Utils.normalize(this.config.headers.sentAt));

    const map = {};
    table.headers.forEach((h, idx) => {
      const key = String(h || '').trim();
      if (!key) return;

      // Exclude reserved/control headers from template variables
      if (reserved.has(Utils.normalize(key))) return;

      map[key] = row.values[idx] == null ? '' : String(row.values[idx]);
    });

    return map;
  }
}
