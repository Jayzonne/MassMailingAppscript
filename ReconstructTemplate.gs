/**
 * ReconstructTemplate.gs
 *
 * Rebuilds the "mass mailing template" layout on the ACTIVE SHEET.
 *
 * Responsibilities:
 * - Reset the active sheet (content + formatting)
 * - Create a standardized header + data layout used by the mass-mailing engine
 * - Apply required merged cells, formatting, and checkbox validations
 * - Populate a minimal "test area" with example values
 *
 * This module is intentionally independent from the sending workflow.
 * It is meant for onboarding, quick resets, and consistent sheet creation.
 */

/**
 * Adds the reconstruction action to an existing custom menu.
 *
 * Usage (inside your single onOpen()):
 *   const menu = ui.createMenu('Send email');
 *   ...
 *   addReconstructTemplateMenu_(menu);
 *   menu.addToUi();
 *
 * @param {GoogleAppsScript.Base.Menu} menu
 *   The menu to which the reconstruction item will be appended.
 */
function addReconstructTemplateMenu_(menu) {
  menu.addSeparator();
  menu.addItem('Reconstruct mass mailing template', 'reconstructMassMailingTemplate');
}

/**
 * Reconstructs the mass mailing template on the currently active sheet.
 *
 * Warning:
 * - This operation clears all content and formatting of the active sheet.
 * - It is destructive by design.
 *
 * Layout produced:
 * - Top configuration area (instructions + template id + subject)
 * - "Test Email Data" row (row 10) with a dedicated checkbox in G10
 * - Header row (row 11) with color-coded sections:
 *   - A..G: email configuration (blue)
 *   - H..L: template variables (green)
 * - Data section starting at row 12 with checkbox validations
 */
function reconstructMassMailingTemplate() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const resp = ui.alert(
    `Reconstruct mass mailing template on sheet "${sheet.getName()}"?\n\nThis will CLEAR the sheet content & formatting.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;

  /**
   * Layout constants.
   * Keep these aligned with the sending engine configuration (APP_CONFIG).
   */
  const HEADER_ROW = 11;
  const FIRST_DATA_ROW = 12;
  const MAX_DATA_ROWS = 300;
  const LAST_COL = 12; // A..L

  /**
   * Column map (1-based indices).
   * This mapping is used only for reconstruction and formatting.
   */
  const COL = {
    TO_SEND: 1,   // A
    SENT: 2,      // B
    EMAIL: 3,     // C
    CC: 4,        // D
    BCC: 5,       // E
    REPLY_TO: 6,  // F
    NO_REPLY: 7,  // G
    NAME: 8,      // H
    TOPIC1: 9,    // I
    TOPIC2: 10,   // J
    TOPIC3: 11,   // K
    COUNTRY: 12,  // L
  };

  // ---------------------------------------------------------------------------
  // Reset sheet
  // ---------------------------------------------------------------------------

  sheet.clear({ contentsOnly: false });
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  // Ensure the sheet has at least A..L columns.
  if (sheet.getMaxColumns() < LAST_COL) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), LAST_COL - sheet.getMaxColumns());
  }

  // ---------------------------------------------------------------------------
  // Layout: column widths (approx. matching the reference sheet)
  // ---------------------------------------------------------------------------

  sheet.setColumnWidth(1, 90);    // A To send
  sheet.setColumnWidth(2, 90);    // B Sent
  sheet.setColumnWidth(3, 260);   // C Email
  sheet.setColumnWidth(4, 200);   // D cc
  sheet.setColumnWidth(5, 200);   // E bcc
  sheet.setColumnWidth(6, 140);   // F replyTo
  sheet.setColumnWidth(7, 110);   // G noReply
  sheet.setColumnWidth(8, 140);   // H Name
  sheet.setColumnWidth(9, 170);   // I Topic1
  sheet.setColumnWidth(10, 140);  // J Topic2
  sheet.setColumnWidth(11, 140);  // K Topic3
  sheet.setColumnWidth(12, 140);  // L Country

  // ---------------------------------------------------------------------------
  // Layout: merged ranges
  // ---------------------------------------------------------------------------
  // Merges must be applied before writing values/formulas into those ranges.
  sheet.getRange('B2:D2').merge();     // instructions text line
  sheet.getRange('B4:C4').merge();     // user guide link placeholder
  sheet.getRange('B6:C6').merge();     // template ID input
  sheet.getRange('F6:G6').merge();     // helper text
  sheet.getRange('B7:D7').merge();     // subject input
  sheet.getRange('A10:B10').merge();   // "Test Email Data" label

  // ---------------------------------------------------------------------------
  // Top area: instructions + configuration cells
  // ---------------------------------------------------------------------------

  sheet.getRange('A2').setValue('Instructions:').setFontWeight('bold');
  sheet.getRange('B2').setValue(
    'This tool allows you to send multiple personalised emails to different recipient'
  );

  sheet.getRange('B4').setValue("=HYPERLINK(\"Link to your tuto\";\"View detailed user guide here\")");


  sheet.getRange('A6').setValue('Template ID:').setFontWeight('bold');

  // Template ID is expected to be pasted by the user.
  sheet.getRange('B6').setValue('1f8xBSdiOR3rbBZIh7eNdZFoaaigu3qI5D3rdMbqiryU');

  // Link to the Google Docs template, built from the template ID cell.
  // Note: Uses French locale separator ";".
  sheet.getRange('E6').setFormula(
    '=HYPERLINK("https://docs.google.com/document/d/" & B6 & "/edit"; "Template")'
  );

  // Helper text shown to the user (merged F6:G6).
  sheet.getRange('F6').setValue('The ID of a Google Doc template');

  sheet.getRange('A7').setValue('Subject:').setFontWeight('bold');
  sheet.getRange('B7').setValue('test email');

  // ---------------------------------------------------------------------------
  // Test area (row 10)
  // ---------------------------------------------------------------------------

  sheet.getRange('A10')
    .setValue('Test Email Data')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // Dedicated checkbox in G10 (test-only control).
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange('G10')
    .setDataValidation(checkboxRule)
    .setValue(false)
    .setHorizontalAlignment('center');

  // Example values (row 10) to demonstrate expected input format.
  sheet.getRange('C10').setValue('example.to@domain.com');
  sheet.getRange('D10').setValue('example.cc@domain.com');
  sheet.getRange('H10').setValue('Jayzonne');
  sheet.getRange('I10').setValue('https://github.com/Jayzonne');

  // ---------------------------------------------------------------------------
  // Header row (row 11)
  // ---------------------------------------------------------------------------

  const headers = [
    'To send', 'Sent', 'Email', 'cc', 'bcc', 'replyTo', 'noReply',
    'Name', 'Topic1', 'Topic2', 'Topic3', 'Country'
  ];

  sheet.getRange(HEADER_ROW, 1, 1, LAST_COL)
    .setValues([headers])
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  // Color-coded header sections (approx. matching the reference sheet).
  const blue = '#4a86e8';
  const green = '#93c47d';

  sheet.getRange(HEADER_ROW, 1, 1, 7).setBackground(blue).setFontColor('#000000');   // A..G
  sheet.getRange(HEADER_ROW, 8, 1, 5).setBackground(green).setFontColor('#000000');  // H..L

  sheet.getRange(HEADER_ROW, 1, 1, LAST_COL)
    .setBorder(true, true, true, true, true, true);

  sheet.setRowHeight(HEADER_ROW, 28);

  // ---------------------------------------------------------------------------
  // Data area (rows 12+): validations + formatting
  // ---------------------------------------------------------------------------

  // Checkboxes for operational columns: To send / Sent / noReply
  sheet.getRange(FIRST_DATA_ROW, COL.TO_SEND, MAX_DATA_ROWS, 1).setDataValidation(checkboxRule);
  sheet.getRange(FIRST_DATA_ROW, COL.SENT, MAX_DATA_ROWS, 1).setDataValidation(checkboxRule);
  sheet.getRange(FIRST_DATA_ROW, COL.NO_REPLY, MAX_DATA_ROWS, 1).setDataValidation(checkboxRule);

  // Center checkbox columns
  sheet.getRange(FIRST_DATA_ROW, COL.TO_SEND, MAX_DATA_ROWS, 2).setHorizontalAlignment('center'); // A,B
  sheet.getRange(FIRST_DATA_ROW, COL.NO_REPLY, MAX_DATA_ROWS, 1).setHorizontalAlignment('center'); // G

  // Force email-related columns to plain text (prevents Sheets from auto-formatting).
  sheet.getRange(FIRST_DATA_ROW, COL.EMAIL, MAX_DATA_ROWS, 1).setNumberFormat('@');
  sheet.getRange(FIRST_DATA_ROW, COL.CC, MAX_DATA_ROWS, 1).setNumberFormat('@');
  sheet.getRange(FIRST_DATA_ROW, COL.BCC, MAX_DATA_ROWS, 1).setNumberFormat('@');
  sheet.getRange(FIRST_DATA_ROW, COL.REPLY_TO, MAX_DATA_ROWS, 1).setNumberFormat('@');

  // Example values (row 12) to demonstrate "real" campaign data format.
  sheet.getRange('C12').setValue('john.doe@domain.com');
  sheet.getRange('D12').setValue('john.doe@domain.com');
  sheet.getRange('H12').setValue('Lorem');
  sheet.getRange('I12').setValue('Truck');

  // Freeze rows through the header so instructions + headers remain visible.
  sheet.setFrozenRows(HEADER_ROW);

  // Small alignment cleanup for the top area.
  sheet.getRange('A2:A7').setHorizontalAlignment('left');
  sheet.getRange('B2').setHorizontalAlignment('left');

  SpreadsheetApp.flush();
}
