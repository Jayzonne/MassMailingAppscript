/**
 * Orchestrator.gs
 *
 * Orchestrates the end-to-end "send" workflow.
 *
 * Design goals:
 * - Keep business flow readable (validate → select → confirm → send → mark → throttle)
 * - Delegate specialized work to services:
 *   - SheetTable: sheet parsing + header lookup
 *   - EmailComposer: Gmail options + template variable mapping
 *   - TemplateRenderer: Google Doc merge + plain text extraction
 *   - MailSender: row marking (Sent/SentAt) + throttling
 *
 * Operational constraints:
 * - Batch sending can hit Gmail / Apps Script quotas; throttling reduces burst behavior.
 * - UI feedback must be explicit: block on structural issues, continue on per-row failures.
 */
class MailOrchestrator {
  /**
   * @param {AppContext} ctx
   *   Runtime context containing active sheet, UI instance, and config.
   */
  constructor(ctx) {
    this.ctx = ctx;

    /**
     * Table snapshot:
     * - We load the sheet once to resolve headers and row values.
     * - If the user edits headers while sending, they must rerun the action.
     */
    this.table = SheetTable.fromSheet(ctx.sheet, ctx.config.headerRow);

    this.renderer = new TemplateRenderer();
    this.composer = new EmailComposer(ctx.config);
    this.sender = new MailSender({ sheet: ctx.sheet, config: ctx.config });
  }

  /**
   * Sends all rows where:
   * - "To send" is checked
   * - "Sent" is not checked
   *
   * Behavior:
   * - Blocks early on structural issues (missing template ID, missing global subject, header errors)
   * - Prompts the user for confirmation before sending
   * - Sends sequentially; marks rows immediately after each successful send
   * - Continues when a row fails; summarizes failures at the end
   */
  sendSelected() {
    const ui = this.ctx.ui;
    const sheet = this.ctx.sheet;
    const cfg = this.ctx.config;

    const templateId = String(sheet.getRange(cfg.templateIdCell).getValue() || '').trim();
    const defaultSubject = String(sheet.getRange(cfg.subjectCell).getValue() || '').trim();

    // These are sheet-level prerequisites; failing them should block any send attempt.
    if (!templateId) return ui.alert(`Missing Template ID in ${cfg.templateIdCell}`);
    if (!defaultSubject) return ui.alert(`Missing Subject in ${cfg.subjectCell} (global subject is required).`);

    /**
     * Duplicate headers create ambiguous mapping (e.g., which "Email" should be used),
     * so we block before sending anything.
     */
    const headerProblems = this.table.validateNoDuplicateHeaders();
    if (headerProblems.length) {
      return ui.alert('Cannot send emails due to header errors:\n\n' + headerProblems.join('\n'));
    }

    /**
     * Minimal header requirements for the selection workflow:
     * - toSend: user selection checkbox
     * - sent: avoids duplicates / resends
     * - email: recipient address
     */
    const idxToSend = this.table.getIndex(cfg.headers.toSend);
    const idxSent = this.table.getIndex(cfg.headers.sent);
    const idxEmail = this.table.getIndex('email');

    if (idxToSend == null || idxSent == null || idxEmail == null) {
      return ui.alert(`Missing required headers. Need: "${cfg.headers.toSend}", "${cfg.headers.sent}", "Email"`);
    }

    /**
     * Selection pass:
     * - We collect candidates and separately collect blocking row-level errors
     *   (e.g., To send checked but Email empty).
     * - We block the whole batch if any selected row is invalid; this prevents partial sends
     *   that might surprise the user.
     */
    const candidates = [];
    const rowErrors = [];

    for (const r of this.table.rows) {
      const wantSend = Utils.asBool(r.values[idxToSend]);
      const alreadySent = Utils.asBool(r.values[idxSent]);
      if (!wantSend || alreadySent) continue;

      const to = Utils.normalizeEmailList(r.values[idxEmail]);
      if (!to) rowErrors.push(`Row ${r.rowNumber}: missing Email`);
      else candidates.push(r);
    }

    if (rowErrors.length) {
      return ui.alert('Fix these rows before sending:\n\n' + rowErrors.slice(0, 30).join('\n'));
    }
    if (!candidates.length) {
      return ui.alert('No rows selected (To send checked) that are not already Sent.');
    }

    /**
     * Confirmation step:
     * - Explicitly show the sheet name to avoid “wrong tab” mistakes.
     * - Explicitly show throttling because it impacts runtime expectations.
     */
    const sheetName = sheet.getName();
    const resp = ui.alert(
      `About to send ${candidates.length} email(s)\n` +
      `From sheet: "${sheetName}"\n\n` +
      `Throttling: ${cfg.throttling.secondsMin}-${cfg.throttling.secondsMax}s.\n\n` +
      `Continue?`,
      ui.ButtonSet.OK_CANCEL
    );
    if (resp !== ui.Button.OK) return;

    const failures = [];
    let sentCount = 0;

    /**
     * Sending loop is sequential by design:
     * - Easier to reason about rate limiting
     * - Safer for immediate row marking
     * - Simplifies troubleshooting (row-by-row)
     */
    for (const r of candidates) {
      // Keeping references scoped per-iteration helps reduce memory pressure in long batches.
      let emailOptions = null;
      let varsMap = null;
      let bodyText = null;

      try {
        emailOptions = this.composer.buildEmailOptions(this.table, r, defaultSubject);
        varsMap = this.composer.buildTemplateVars(this.table, r);
        bodyText = this.renderer.renderDocToText(templateId, varsMap);

        GmailApp.sendEmail(emailOptions.to, emailOptions.subject, bodyText, emailOptions.options);

        sentCount++;

        // Row marking is done immediately so the UI reflects progress in real time.
        if (cfg.marking.markSentImmediately) {
          this.sender.markSentNow(this.table, r.rowNumber);
        }
      } catch (e) {
        failures.push(`Row ${r.rowNumber}: ${e && e.message ? e.message : String(e)}`);
      } finally {
        emailOptions = null;
        varsMap = null;
        bodyText = null;

        // Throttle regardless of success/failure to keep pacing stable.
        this.sender.throttle();
      }
    }

    if (failures.length) {
      return ui.alert(
        `Done with errors.\n\nSent: ${sentCount}\nFailed: ${failures.length}\n\n` +
        failures.slice(0, 30).join('\n')
      );
    }

    return ui.alert(`All done. Sent ${sentCount} email(s).`);
  }

  /**
   * Sends a single “test” email using an absolute row number.
   *
   * Use case:
   * - Test row lives above the header (e.g., row 10 "Test Email Data")
   * - Allows validating template variables and mail parameters without touching campaign rows
   *
   * @param {number} rowNumber
   *   Absolute 1-based sheet row number.
   */
  sendTestRow(rowNumber) {
    const ui = this.ctx.ui;
    const sheet = this.ctx.sheet;
    const cfg = this.ctx.config;

    const templateId = String(sheet.getRange(cfg.templateIdCell).getValue() || '').trim();
    const defaultSubject = String(sheet.getRange(cfg.subjectCell).getValue() || '').trim();

    if (!templateId) return ui.alert(`Missing Template ID in ${cfg.templateIdCell}`);
    if (!defaultSubject) return ui.alert(`Missing Subject in ${cfg.subjectCell} (global subject is required).`);

    const headerProblems = this.table.validateNoDuplicateHeaders();
    if (headerProblems.length) {
      return ui.alert('Cannot send test email due to header errors:\n\n' + headerProblems.join('\n'));
    }

    const r = this.table.readAbsoluteRow(rowNumber);

    /**
     * Test sending still relies on the header mapping for the "Email" column.
     * If the header is missing, we cannot interpret rowNumber reliably.
     */
    const idxEmail = this.table.getIndex('email');
    const to = idxEmail != null ? Utils.normalizeEmailList(r.values[idxEmail]) : '';

    if (idxEmail == null) return ui.alert(`Cannot find "Email" header on row ${cfg.headerRow}.`);
    if (!to) return ui.alert(`Row ${rowNumber}: "Email" is empty.`);

    try {
      const emailOptions = this.composer.buildEmailOptions(this.table, r, defaultSubject);
      const varsMap = this.composer.buildTemplateVars(this.table, r);
      const bodyText = this.renderer.renderDocToText(templateId, varsMap);

      GmailApp.sendEmail(emailOptions.to, emailOptions.subject, bodyText, emailOptions.options);

      return ui.alert(`Test email sent using row ${rowNumber}.`);
    } catch (e) {
      return ui.alert(`Test email failed for row ${rowNumber}:\n\n${e && e.message ? e.message : String(e)}`);
    }
  }
}
