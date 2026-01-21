/**
 * AppContext.gs
 *
 * Application context object used to share
 * runtime dependencies across the system.
 *
 * This class centralizes:
 * - the active Google Sheet
 * - the Google Sheets UI instance
 * - the global application configuration
 *
 * It acts as a lightweight dependency container,
 * allowing services and orchestrators to remain
 * decoupled from global APIs.
 */
class AppContext {

  /**
   * Creates a new application context.
   *
   * @param {Object} params
   * @param {GoogleAppsScript.Spreadsheet.Sheet} params.sheet
   *   The active sheet used as the source of the mailing data.
   * @param {GoogleAppsScript.Base.Ui} params.ui
   *   UI instance used to display alerts and confirmations.
   * @param {Object} params.config
   *   Central application configuration (APP_CONFIG).
   */
  constructor({ sheet, ui, config }) {
    this.sheet = sheet;
    this.ui = ui;
    this.config = config;
  }

  /**
   * Factory method creating an AppContext
   * from the currently active Google Sheet.
   *
   * This is the default entry point used by
   * UI actions (menu items, buttons).
   *
   * @returns {AppContext}
   */
  static fromActiveSheet_() {
    return new AppContext({
      sheet: SpreadsheetApp.getActiveSheet(),
      ui: SpreadsheetApp.getUi(),
      config: APP_CONFIG,
    });
  }
}
