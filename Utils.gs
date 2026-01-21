/**
 * Utils.gs
 *
 * Collection of stateless utility helpers shared across the application.
 *
 * This class contains only pure, side-effect-free helpers
 * (except for controlled sleep used for throttling).
 *
 * Utilities are intentionally grouped here to:
 * - avoid duplication
 * - keep business logic readable
 * - centralize low-level transformations
 */
class Utils {

  /**
   * Normalizes a value for reliable comparisons.
   *
   * - Converts to string
   * - Trims leading/trailing whitespace
   * - Lowercases the result
   *
   * Commonly used for:
   * - header name matching
   * - configuration keys
   * - defensive string comparisons
   *
   * @param {*} s
   * @returns {string}
   */
  static normalize(s) {
    return String(s || '').trim().toLowerCase();
  }

  /**
   * Converts a spreadsheet cell value into a boolean.
   *
   * Truthy values:
   * - true (boolean)
   * - "true", "yes", "1" (case-insensitive)
   *
   * All other values are considered false.
   *
   * @param {*} v
   * @returns {boolean}
   */
  static asBool(v) {
    if (v === true) return true;
    const s = String(v || '').trim().toLowerCase();
    return s === 'true' || s === 'yes' || s === '1';
  }

  /**
   * Escapes a string so it can be safely injected
   * into a regular expression pattern.
   *
   * This is required when replacing template variables
   * whose names may contain special regex characters.
   *
   * @param {string} s
   * @returns {string}
   */
  static escapeRegex(s) {
    return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  /**
   * Normalizes a list of email addresses from a spreadsheet cell.
   *
   * Supported input formats:
   * - "a@x.com, b@y.com"
   * - "a@x.com;b@y.com"
   * - multiline values
   *
   * Output format:
   * - "a@x.com, b@y.com"
   *
   * Empty values are removed.
   *
   * @param {*} s
   * @returns {string}
   */
  static normalizeEmailList(s) {
    return String(s || '')
      .replace(/\n/g, ' ')
      .replace(/;/g, ',')
      .split(',')
      .map(x => x.trim())
      .filter(Boolean)
      .join(', ');
  }

  /**
   * Sleeps for a random duration between two bounds (in seconds).
   *
   * Used to:
   * - throttle email sending
   * - reduce Gmail anti-spam triggers
   * - avoid Apps Script rate limits
   *
   * The randomness prevents predictable burst patterns.
   *
   * @param {number} minSeconds
   * @param {number} maxSeconds
   */
  static sleepRandomSeconds(minSeconds, maxSeconds) {
    const minMs = Math.max(0, (minSeconds || 0) * 1000);
    const maxMs = Math.max(minMs, (maxSeconds || 0) * 1000);

    const ms = (minMs === maxMs)
      ? minMs
      : (minMs + Math.floor(Math.random() * (maxMs - minMs + 1)));

    Utilities.sleep(ms);
  }
}
