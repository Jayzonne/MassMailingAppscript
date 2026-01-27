/**
 * Services_EmailComposer.gs
 *
 * Builds two things from a single sheet row:
 * - Gmail send arguments (recipient list, subject, Gmail options)
 * - Template variable map used by the Google Docs mail-merge renderer
 *
 * Key rules:
 * - Email parameters must come from "reserved" (blue) headers only (Email, cc, bcc, replyTo, noReply, etc.)
 * - Template variables are derived from all non-reserved headers (green section + any custom fields)
 * - Subject resolution: row-level "Subject" overrides the global subject; global is required upstream
 *
 * This service is pure preparation logic: it does not send emails and does not write to the sheet.
 */
class EmailComposer {
  /**
   * @param {Object} config
   *   APP_CONFIG used for:
   *   - reserved header names
   *   - no-reply sender address
   */
  constructor(config) {
    this.config = config;
  }

  /**
   * Builds the Gmail parameters for a row.
   *
   * Subject resolution:
   * - If the row contains a "Subject" value, it wins.
   * - Otherwise, the global subject is used (and is validated by the orchestrator).
   *
   * Address fields (Email / cc / bcc / replyTo):
   * - Accept multiple recipients in a single cell.
   * - Normalization supports comma-separated lists and converts ";" to ",".
   *
   * noReply behavior:
   * - When enabled, the sender is forced to config.noReplyFromEmail.
   * - That address must be configured as an allowed alias in the account ("Send mail as").
   *
   * @param {SheetTable} table
   *   Parsed table with header indexing.
   * @param {{rowNumber:number, values:any[]}} row
   *   Source row values.
   * @param {string} defaultSubject
   *   Global subject fallback (required by the orchestrator).
   * @returns {{to:string, subject:string, options:Object}}
   *   Arguments suitable for GmailApp.sendEmail(to, subject, body, options).
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
   * Builds the variable map used for template rendering.
   *
   * Rule:
   * - Every header that is NOT reserved becomes a template variable key.
   * - Reserved headers include both:
   *   - email-parameter headers (Email, cc, bcc, replyTo, noReply, subject, etc.)
   *   - control/status headers (to send, sent, sentAt)
   *
   * Placeholder format expected by TemplateRenderer:
   * - Sheet header: Topic1
   * - Google Doc placeholder: $Topic1$ (or $ Topic1 $)
   *
   * @param {SheetTable} table
   * @param {{rowNumber:number, values:any[]}} row
   * @returns {Object.<string,string>}
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

      if (reserved.has(Utils.normalize(key))) return;

      map[key] = row.values[idx] == null ? '' : String(row.values[idx]);
    });

    return map;
  }
}
