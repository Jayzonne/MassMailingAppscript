/**
 * Orchestrator.gs
 *
 * MailOrchestrator coordinates the end-to-end workflow for the mass mailing system.
 *
 * Responsibilities:
 * - Read sheet-level configuration (template ID, default subject)
 * - Validate table structure (headers, required columns)
 * - Identify which rows should be sent (To send checked, Sent unchecked)
 * - For each row:
 *   - Compose Gmail options (EmailComposer)
 *   - Build template variables (EmailComposer)
 *   - Render the Google Docs template into text (TemplateRenderer)
 *   - Send the email (GmailApp)
 *   - Mark the row as sent immediately (MailSender)
 *   - Apply throttling between attempts (MailSender)
 *
 * This class is the "business workflow layer":
 * - It contains orchestration logic and user-facing validations
 * - It delegates specialized tasks to services
 * - It avoids low-level spreadsheet parsing and templating details
 */
class MailOrchestrator {

  /**
   * Creates a new MailOrchestrator.
   *
   * Dependencies are constructed here to keep the UI layer minimal.
   *
   * @param {AppContext} ctx
   *   Runtime context containing the active sheet, UI instance, and config.
   */
  constructor(ctx) {
    this.ctx = ctx;

    // Table snapshot is captured at construction time.
    // If headers change during execution, recreate the orchestrator.
    this.table = SheetTable.fromSheet(ctx.sheet, ctx.config.headerRow);

    this.renderer = new TemplateRenderer();
    this.composer = new EmailComposer(ctx.config);
    this.sender = new MailSender({ sheet: ctx.sheet, config: ctx.config });
  }

  /**
   * Sends emails for all rows that are selected for sending.
   *
   * Selection rule:
   * - "To send" is checked
   * - "Sent" is not checked
   *
   * Safety checks:
   * - Template ID and Subject must be present
   * - Headers must be unique
   * - Required headers must exist
   * - Selected rows must have a valid recipient ("Email")
   *
   * Operational behavior:
   * - Sends emails one by one with throttling
   * - Marks each row as sent immediately after success
   * - Continues after failures (failures are aggregated and shown at the end)
   */
  sendSelected() {
    const ui = this.ctx.ui;
    const sheet = this.ctx.sheet;
    const cfg = this.ctx.config;

    // Sheet-level configuration inputs
    const templateId = String(sheet.getRange(cfg.templateIdCell).getValue() || '').trim();
    const defaultSubject = String(sheet.getRange(cfg.subjectCell).getValue() || '').trim();

    if (!templateId) return ui.alert(`Missing Template ID in ${cfg.templateIdCell}`);
    if (!defaultSubject) return ui.alert(`Missing Subject in ${cfg.subjectCell}`);

    // Structural validation
    const headerProblems = this.table.validateNoDuplicateHeaders();
    if (headerProblems.length) {
      return ui.alert('Cannot send emails due to header errors:\n\n' + headerProblems.join('\n'));
    }

    // Required columns
    const idxToSend = this.table.getIndex(cfg.headers.toSend);
    const idxSent = this.table.getIndex(cfg.headers.sent);
    const idxEmail = this.table.getIndex('email');

    if (idxToSend == null || idxSent == null || idxEmail == null) {
      return ui.alert(
        `Missing required headers. Need: "${cfg.headers.toSend}", "${cfg.headers.sent}", "Email"`
      );
    }

    // Identify candidate rows to send
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

    // User confirmation with clear source sheet identification
    const sheetName = sheet.getName();
    const resp = ui.alert(
      `About to send ${candidates.length} email(s)\n` +
      `From sheet: "${sheetName}"\n\n` +
      `Throttling: ${cfg.throttling.secondsMin}-${cfg.throttling.secondsMax}s.\n\n` +
      `Continue?`,
      ui.ButtonSet.OK_CANCEL
    );
    if (resp !== ui.Button.OK) return;

    // Batch execution (failure-tolerant)
    const failures = [];
    let sentCount = 0;

    for (const r of candidates) {
      // Keep object lifetimes small to reduce memory pressure during long batches
      let emailOptions = null;
      let varsMap = null;
      let bodyText = null;

      try {
        // Prepare the email
        emailOptions = this.composer.buildEmailOptions(this.table, r, defaultSubject);

        // Prepare template data and render final body text
        varsMap = this.composer.buildTemplateVars(this.table, r);
        bodyText = this.renderer.renderDocToText(templateId, varsMap);

        // Send via Gmail
        GmailApp.sendEmail(emailOptions.to, emailOptions.subject, bodyText, emailOptions.options);

        // Post-send updates
        sentCount++;
        if (cfg.marking.markSentImmediately) {
          this.sender.markSentNow(this.table, r.rowNumber);
        }

      } catch (e) {
        // Failures are captured and shown at the end; execution continues for other rows
        failures.push(`Row ${r.rowNumber}: ${e && e.message ? e.message : String(e)}`);

      } finally {
        // Best-effort memory cleanup (Apps Script does not expose manual GC)
        emailOptions = null;
        varsMap = null;
        bodyText = null;

        // Controlled pacing to reduce rate-limit and anti-spam issues
        this.sender.throttle();
      }
    }

    // Final status summary
    if (failures.length) {
      return ui.alert(
        `Done with errors.\n\nSent: ${sentCount}\nFailed: ${failures.length}\n\n` +
        failures.slice(0, 30).join('\n')
      );
    }
    return ui.alert(`All done. Sent ${sentCount} email(s).`);
  }

  /**
   * Sends a single test email using an absolute sheet row number.
   *
   * This is designed for the "Test Email Data" area which may live
   * above the table header row, and therefore is not part of the batch rows.
   *
   * Safety checks:
   * - Template ID and Subject must exist
   * - Headers must be unique
   * - The "Email" header must exist
   * - The target row must contain a non-empty "Email" value
   *
   * @param {number} rowNumber
   *   Absolute 1-based row number to send as a test.
   */
  sendTestRow(rowNumber) {
    const ui = this.ctx.ui;
    const sheet = this.ctx.sheet;
    const cfg = this.ctx.config;

    const templateId = String(sheet.getRange(cfg.templateIdCell).getValue() || '').trim();
    const defaultSubject = String(sheet.getRange(cfg.subjectCell).getValue() || '').trim();

    if (!templateId) return ui.alert(`Missing Template ID in ${cfg.templateIdCell}`);
    if (!defaultSubject) return ui.alert(`Missing Subject in ${cfg.subjectCell}`);

    const headerProblems = this.table.validateNoDuplicateHeaders();
    if (headerProblems.length) {
      return ui.alert('Cannot send test email due to header errors:\n\n' + headerProblems.join('\n'));
    }

    // Read row outside of table data area (absolute row read)
    const r = this.table.readAbsoluteRow(rowNumber);

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
      return ui.alert(
        `Test email failed for row ${rowNumber}:\n\n${e && e.message ? e.message : String(e)}`
      );
    }
  }
}
