/**
 * Rebuilds the mass-mailing sheet layout on the active sheet.
 *
 * Why this exists:
 * - People copy/paste or tweak sheets and slowly break the expected structure.
 * - This function provides a one-click, deterministic way to return to a known-good layout.
 *
 * Contract with the rest of the project:
 * - The header row and header labels produced here must match what the sending engine expects
 *   (SheetTable + EmailComposer + Orchestrator).
 * - Changing column order or header names here requires updating the engine configuration accordingly.
 *
 * Notes:
 * - This function intentionally does a "hard reset" (break merges, clear values, clear validations)
 *   to avoid leftover artifacts when re-running reconstruction on a previously edited sheet.
 */
function reconstructMassMailingTemplate() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const resp = ui.alert(
    `Reconstruct mass mailing template on sheet "${sheet.getName()}"?\n\nThis will CLEAR the entire sheet.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;

  /**
   * Layout anchors.
   * - HEADER_ROW must stay aligned with APP_CONFIG.headerRow.
   * - FIRST_DATA_ROW must be HEADER_ROW + 1.
   */
  const HEADER_ROW = 11;
  const FIRST_DATA_ROW = 12;
  const MAX_DATA_ROWS = 300;

  /**
   * This template currently uses 14 columns (A..N).
   * If you add more fields, update:
   * - LAST_COL
   * - COL mapping
   * - headers array
   * - header coloring ranges (blue/green segments)
   */
  const LAST_COL = 14;

  /**
   * Column mapping (1-based).
   * This is used to keep the code resilient if the column order evolves.
   */
  const COL = {
    TO_SEND: 1,
    SENT: 2,
    SENT_AT: 3,
    SUBJECT: 4,
    EMAIL: 5,
    CC: 6,
    BCC: 7,
    REPLY_TO: 8,
    NO_REPLY: 9,
    NAME: 10,
    TOPIC1: 11,
    TOPIC2: 12,
    TOPIC3: 13,
    COUNTRY: 14,
  };

  /**
   * Default values used to make a freshly reconstructed sheet immediately usable.
   * - templateIdDefault is pre-filled to reduce onboarding friction.
   * - globalSubjectDefault is required by the engine and blocks sending if empty.
   * - userGuideUrl is a clickable entry point for end users.
   */
  const templateIdDefault = '1f8xBSdiOR3rbBZIh7eNdZFoaaigu3qI5D3rdMbqiryU';
  const globalSubjectDefault = 'Pour me soutenir => une étoile sur mon GitHub.';
  const userGuideUrl = 'https://github.com/jayzonne';

  /**
   * Hard reset: ensures reconstruction is idempotent.
   * We break merges before clearing to avoid residual merge artifacts.
   */
  const fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  fullRange.breakApart();
  sheet.clear();
  fullRange.clearDataValidations();
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  sheet.setRowHeights(1, sheet.getMaxRows(), 21);

  // Ensure enough columns exist for the template (A..N).
  if (sheet.getMaxColumns() < LAST_COL) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), LAST_COL - sheet.getMaxColumns());
  }

  /**
   * Column sizing is UX-driven.
   * - Column B ("Sent") is the baseline reference.
   * - A and C are slightly wider for checkbox clarity + full timestamp readability.
   */
  sheet.setColumnWidth(COL.TO_SEND, 130);
  sheet.setColumnWidth(COL.SENT, 90);
  sheet.setColumnWidth(COL.SENT_AT, 235);
  sheet.setColumnWidth(COL.SUBJECT, 300);
  sheet.setColumnWidth(COL.EMAIL, 280);
  sheet.setColumnWidth(COL.CC, 200);
  sheet.setColumnWidth(COL.BCC, 200);
  sheet.setColumnWidth(COL.REPLY_TO, 160);
  sheet.setColumnWidth(COL.NO_REPLY, 110);
  sheet.setColumnWidth(COL.NAME, 140);
  sheet.setColumnWidth(COL.TOPIC1, 220);
  sheet.setColumnWidth(COL.TOPIC2, 140);
  sheet.setColumnWidth(COL.TOPIC3, 140);
  sheet.setColumnWidth(COL.COUNTRY, 140);

  /**
   * Merges are part of the "template UX" (top configuration zone).
   * If you move config cells (e.g., Template ID), update APP_CONFIG accordingly.
   */
  sheet.getRange('B2:D2').merge();
  sheet.getRange('B4:C4').merge();
  sheet.getRange('B6:C6').merge();
  sheet.getRange('E6:F6').merge();
  sheet.getRange('B7:D7').merge();
  sheet.getRange('A10:B10').merge();

  sheet.getRange('A2').setValue('Instructions:').setFontWeight('bold');
  sheet.getRange('B2').setValue('This tool allows you to send multiple personalised emails to different recipient');

  // Use HYPERLINK formula for portability across copies of the sheet.
  sheet.getRange('B4').setFormula(`=HYPERLINK("${userGuideUrl}"; "View detailed user guide here")`);

  sheet.getRange('A6').setValue('Template ID:').setFontWeight('bold');
  sheet.getRange('B6').setValue(templateIdDefault);

  /**
   * The "Template" link is derived from the template ID cell.
   * Locale note: using ";" as argument separator (French locale sheets).
   * If your spreadsheet locale uses ",", update the formula separator.
   */
  sheet.getRange('E6').setFormula(
    '=HYPERLINK("https://docs.google.com/document/d/" & B6 & "/edit"; "Template")'
  );
  sheet.getRange('G6').setValue('The ID of a Google Doc template');

  // Global subject is required: the sending workflow blocks if this is empty.
  sheet.getRange('A7').setValue('Subject (default):').setFontWeight('bold');
  sheet.getRange('B7').setValue(globalSubjectDefault);

  sheet.getRange('A10').setValue('Test Email Data').setFontWeight('bold');

  /**
   * Checkboxes are used as an explicit, low-friction workflow:
   * - To send: user intent selection
   * - Sent: system status
   * - noReply: per-row sending behavior
   */
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet
    .getRange(10, COL.NO_REPLY)
    .setDataValidation(checkboxRule)
    .setValue(false)
    .setHorizontalAlignment('center');

  /**
   * Test row sample values are intentionally recognizable,
   * making it easy to validate template variables end-to-end.
   */
  sheet.getRange(10, COL.SUBJECT).setValue('Test subject');
  sheet.getRange(10, COL.EMAIL).setValue('example.to@domain.com');
  sheet.getRange(10, COL.CC).setValue('example.cc@domain.com');
  sheet.getRange(10, COL.NAME).setValue('Jayzonne');
  sheet.getRange(10, COL.TOPIC1).setValue('github.com/jayzonne');

  /**
   * Header names are the API between the sheet and the engine.
   * Changes here must be reflected in:
   * - APP_CONFIG.headers
   * - APP_CONFIG.reservedEmailHeaders
   * - any logic expecting these labels (EmailComposer, Orchestrator)
   */
  const headers = [
    'To send',
    'Sent',
    'SentAt',
    'Subject',
    'Email',
    'cc',
    'bcc',
    'replyTo',
    'noReply',
    'Name',
    'Topic1',
    'Topic2',
    'Topic3',
    'Country',
  ];

  sheet
    .getRange(HEADER_ROW, 1, 1, LAST_COL)
    .setValues([headers])
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  /**
   * Visual separation:
   * - Blue: email configuration / operational columns
   * - Green: template variables (mail merge payload)
   *
   * Keep these ranges aligned with the column split.
   */
  const blue = '#4a86e8';
  const green = '#93c47d';

  sheet.getRange(HEADER_ROW, 1, 1, 9).setBackground(blue).setFontColor('#000000');
  sheet.getRange(HEADER_ROW, 10, 1, 5).setBackground(green).setFontColor('#000000');
  sheet.getRange(HEADER_ROW, 1, 1, LAST_COL).setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(HEADER_ROW, 28);

  sheet.getRange(FIRST_DATA_ROW, COL.TO_SEND, MAX_DATA_ROWS, 1).setDataValidation(checkboxRule);
  sheet.getRange(FIRST_DATA_ROW, COL.SENT, MAX_DATA_ROWS, 1).setDataValidation(checkboxRule);
  sheet.getRange(FIRST_DATA_ROW, COL.NO_REPLY, MAX_DATA_ROWS, 1).setDataValidation(checkboxRule);

  sheet.getRange(FIRST_DATA_ROW, COL.TO_SEND, MAX_DATA_ROWS, 2).setHorizontalAlignment('center');
  sheet.getRange(FIRST_DATA_ROW, COL.NO_REPLY, MAX_DATA_ROWS, 1).setHorizontalAlignment('center');

  /**
   * Timestamp formatting:
   * - Uses APP_CONFIG.sentAtNumberFormat if present to keep a single source of truth.
   * - Center alignment reinforces "system-generated / do not edit".
   */
  sheet
    .getRange(FIRST_DATA_ROW, COL.SENT_AT, MAX_DATA_ROWS, 1)
    .setNumberFormat((APP_CONFIG && APP_CONFIG.sentAtNumberFormat) ? APP_CONFIG.sentAtNumberFormat : 'yyyy/MM/dd - HH:mm:ss')
    .setHorizontalAlignment('center');

  /**
   * Force email-like fields to plain text to prevent Sheets auto-formatting.
   * This avoids issues like stripping leading characters or converting to links unexpectedly.
   */
  sheet.getRange(FIRST_DATA_ROW, COL.SUBJECT, MAX_DATA_ROWS, 1).setNumberFormat('@');
  sheet.getRange(FIRST_DATA_ROW, COL.EMAIL, MAX_DATA_ROWS, 1).setNumberFormat('@');
  sheet.getRange(FIRST_DATA_ROW, COL.CC, MAX_DATA_ROWS, 1).setNumberFormat('@');
  sheet.getRange(FIRST_DATA_ROW, COL.BCC, MAX_DATA_ROWS, 1).setNumberFormat('@');
  sheet.getRange(FIRST_DATA_ROW, COL.REPLY_TO, MAX_DATA_ROWS, 1).setNumberFormat('@');

  /**
   * Read-only hinting:
   * Sent + SentAt are system-maintained columns; shading discourages manual edits.
   * (If you want hard enforcement, add sheet protections in a dedicated helper.)
   */
  const readOnlyGray = '#eeeeee';
  sheet.getRange(FIRST_DATA_ROW, COL.SENT, MAX_DATA_ROWS, 1).setBackground(readOnlyGray);
  sheet.getRange(FIRST_DATA_ROW, COL.SENT_AT, MAX_DATA_ROWS, 1).setBackground(readOnlyGray);

  sheet.setFrozenRows(HEADER_ROW);
  SpreadsheetApp.flush();
}
