/**
 * Config.gs
 *
 * Central configuration for the Google Sheets Mass Mailing Engine.
 *
 * Purpose:
 * - Define all sheet-specific and behavior-specific settings in one place
 * - Allow adapting the system to a new sheet layout without touching logic code
 * - Act as a contract between the spreadsheet structure and the engine
 *
 * Any change here should be intentional and reviewed, as it can affect:
 * - Header resolution
 * - Sending eligibility
 * - Throttling behavior
 * - Row state management
 */
const APP_CONFIG = {
  /**
   * Row number containing the column headers.
   *
   * Must match the value used by:
   * - SheetTable.fromSheet(...)
   * - Template reconstruction logic
   */
  headerRow: 11,

  /**
   * Cell containing the Google Docs template ID.
   * Used by the orchestrator to locate the source document for mail-merge.
   */
  templateIdCell: 'B6',

  /**
   * Cell containing the global (default) subject.
   *
   * Rule:
   * - This value is required to send emails
   * - A row-level "Subject" overrides this value when present
   */
  subjectCell: 'B7',

  /**
   * Canonical names for control/status headers.
   *
   * These are normalized internally (case-insensitive),
   * but must remain stable across the system.
   */
  headers: {
    toSend: 'to send',
    sent: 'sent',
    sentAt: 'sentAt',
  },

  /**
   * Throttling configuration (in seconds).
   *
   * Randomized delays between sends are used to:
   * - Reduce burst-like behavior
   * - Lower the risk of hitting Gmail or Apps Script quotas
   */
  throttling: {
    secondsMin: 10,
    secondsMax: 15,
  },

  /**
   * Row-marking behavior after a successful send.
   *
   * These flags allow adjusting UX without changing orchestration logic.
   */
  marking: {
    markSentImmediately: true,      // Update status as soon as an email is sent
    clearToSendAfterSent: true,     // Prevent accidental re-sends
    writeSentTimestamp: true,       // Enable sentAt date/time tracking
  },

  /**
   * Configuration for the "test email" feature.
   *
   * The test row is intentionally outside the main data block
   * to avoid accidental inclusion in batch sends.
   */
  test: {
    rowNumber: 10,
  },

  /**
   * Reserved headers are excluded from template variable substitution.
   *
   * Rationale:
   * - Email routing and control fields must never leak into template variables
   * - Variants are included to tolerate common header spelling differences
   *
   * If you add a new system-level column, it should be listed here.
   */
  reservedEmailHeaders: [
    'email',
    'cc',
    'bcc',
    'reply to',
    'replyto',
    'reply_to',
    'no reply',
    'noreply',
    'no_reply',
    'fromname',
    'from name',
    'fromemail',
    'from email',
    'subject',
    'sentat',
    'sent at',
    'sent_at',
  ],

  /**
   * Sender address used when the "noReply" flag is enabled on a row.
   *
   * Important:
   * - This address must be configured as an allowed alias
   *   in the Gmail account ("Send mail as").
   */
  noReplyFromEmail: 'noreply@domain.com',

  /**
   * Canonical date/time format for the SentAt column.
   *
   * Centralizing this here ensures:
   * - Consistent rendering between reconstruction and runtime updates
   * - Easy localization or format changes in one place
   */
  sentAtNumberFormat: 'yyyy/MM/dd - HH:mm:ss',
};
