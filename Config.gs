/**
 * Config.gs
 *
 * Central configuration for the Google Sheets Mass Mailing Engine.
 *
 * This file defines:
 * - Structural rules of the spreadsheet (header row, fixed cells)
 * - Email-sending behavior (throttling, marking strategy)
 * - Reserved column names used for email configuration
 * - Test mode configuration
 *
 * ⚠️ This file should be considered read-only for end users.
 * Any change here impacts the global behavior of the system.
 */
const APP_CONFIG = {

  /**
   * Row index (1-based) where the table headers are located.
   * All data rows must be strictly below this row.
   */
  headerRow: 11,

  /**
   * Cell containing the Google Docs template ID.
   * Example: B6
   */
  templateIdCell: 'B6',

  /**
   * Cell containing the default email subject.
   * Can be overridden per row if a "subject" column exists.
   */
  subjectCell: 'B7',

  /**
   * Canonical header names used internally by the system.
   * These values must match the column headers in the sheet
   * (case-insensitive, whitespace-insensitive).
   */
  headers: {
    /** Checkbox column indicating a row should be sent */
    toSend: 'to send',

    /** Checkbox column indicating the email has been successfully sent */
    sent: 'sent',

    /** Optional column storing the timestamp of the send operation */
    sentAt: 'sentAt',
  },

  /**
   * Throttling configuration to avoid Gmail anti-spam
   * and Apps Script rate limits.
   *
   * A random delay between min and max is applied
   * after each email send attempt.
   */
  throttling: {
    /** Minimum delay (in seconds) between two emails */
    secondsMin: 10,

    /** Maximum delay (in seconds) between two emails */
    secondsMax: 15,
  },

  /**
   * Behavior related to how rows are updated after sending.
   */
  marking: {
    /**
     * If true, the "Sent" checkbox is marked immediately
     * after each successful email (not at end of batch).
     */
    markSentImmediately: true,

    /**
     * If true, the "To send" checkbox is cleared
     * once the email has been sent.
     */
    clearToSendAfterSent: true,

    /**
     * If true and a "sentAt" column exists,
     * the current timestamp is written after sending.
     */
    writeSentTimestamp: true,
  },

  /**
   * Configuration for the built-in test mode.
   * This row is sent independently of the batch logic.
   */
  test: {
    /**
     * Absolute row number used for "Test email" action.
     * This row may be above the header row.
     */
    rowNumber: 10,
  },

  /**
   * List of reserved column headers used for email configuration.
   *
   * These columns are considered "blue columns" and:
   * - are used to build Gmail options (To, CC, BCC, etc.)
   * - are excluded from Google Docs template variables
   *
   * Any duplicate or misuse of these headers will block sending.
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
    'subject'
  ],

  /**
   * Email address used when "No Reply" is enabled on a row.
   *
   * ⚠️ This address MUST be a valid "Send mail as" alias
   * configured in the Gmail / Google Workspace account.
   *
   * Spoofing is intentionally not supported.
   */
  noReplyFromEmail: 'noreply@domain.com',
};
