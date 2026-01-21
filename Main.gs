/**
 * Main.gs
 *
 * Entry points for the Google Sheets UI integration.
 *
 * This file contains:
 * - the `onOpen()` trigger that registers the custom menu in Google Sheets
 * - lightweight handlers that delegate work to the orchestrator layer
 *
 * No business logic should live here.
 * Keep this file limited to UI wiring and high-level routing.
 */

/**
 * Simple trigger executed when the spreadsheet is opened.
 * Registers the "Send email" menu in the Google Sheets UI.
 *
 * Menu actions:
 * - Send selected emails: sends rows where "To send" is checked and "Sent" is not checked
 * - Test email: sends a single configured test row (may be above the header row)
 * - Reconstruct mass mailing template: rebuilds the sheet layout from scratch
 *
 * Note:
 * - Only one onOpen() should exist in the project. Merge menus here if needed.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Send email');

  menu.addItem('Send selected emails', 'sendSelectedEmails');
  menu.addItem(`Test email (send row ${APP_CONFIG.test.rowNumber})`, 'sendTestEmail');

  menu.addSeparator();
  menu.addItem('Reconstruct mass mailing template', 'reconstructMassMailingTemplate');

  menu.addToUi();
}

/**
 * Sends emails for the currently active sheet.
 *
 * The active sheet is treated as the source table:
 * - reads header row and data rows
 * - renders the template per row
 * - sends emails with throttling and immediate "Sent" marking
 */
function sendSelectedEmails() {
  const ctx = AppContext.fromActiveSheet_();
  new MailOrchestrator(ctx).sendSelected();
}

/**
 * Sends a single test email using the configured test row number.
 *
 * This action bypasses batch selection logic and is meant for:
 * - template validation
 * - sender validation (from / replyTo / noReply behavior)
 * - quick end-to-end checks without touching the campaign rows
 */
function sendTestEmail() {
  const ctx = AppContext.fromActiveSheet_();
  new MailOrchestrator(ctx).sendTestRow(APP_CONFIG.test.rowNumber);
}
