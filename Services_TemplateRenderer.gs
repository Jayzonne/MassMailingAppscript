/**
 * Services_TemplateRenderer.gs
 *
 * Renders a Google Docs template into plain text by performing a mail-merge:
 * - Copies the template (so the original is never mutated)
 * - Replaces placeholders of the form $Variable$ (whitespace tolerant)
 * - Returns the final body as plain text
 *
 * Placeholder rules:
 * - Variable names come from sheet headers (case-sensitive on the sheet side, matched literally here)
 * - In the Doc, both `$Topic1$` and `$ Topic1 $` are supported
 *
 * Cleanup:
 * - The temporary copy is always trashed, even if replacement or reading fails
 */
class TemplateRenderer {
  /**
   * @param {string} templateId
   *   Google Docs file ID used as the mail-merge template.
   * @param {Object.<string, string>} varsMap
   *   Map of variable names to replacement values.
   * @returns {string}
   *   Rendered document body as plain text.
   */
  renderDocToText(templateId, varsMap) {
    const file = DriveApp.getFileById(templateId);
    const copy = file.makeCopy(`tmp-mailmerge-${Date.now()}`);

    try {
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();

      /**
       * We escape variable keys to avoid accidental regex metacharacter interpretation.
       * The replacement pattern tolerates whitespace inside the $...$ wrapper.
       */
      Object.keys(varsMap).forEach((key) => {
        const value = varsMap[key] == null ? '' : String(varsMap[key]);
        const escapedKey = Utils.escapeRegex(key);
        const pattern = `\\$\\s*${escapedKey}\\s*\\$`;
        body.replaceText(pattern, value);
      });

      doc.saveAndClose();

      /**
       * Re-opening after save ensures the returned text reflects the committed document state,
       * which avoids edge cases where a cached DocumentApp instance returns stale content.
       */
      return DocumentApp.openById(copy.getId()).getBody().getText();
    } finally {
      copy.setTrashed(true);
    }
  }
}
