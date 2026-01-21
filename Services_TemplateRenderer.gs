/**
 * Services_TemplateRenderer.gs
 *
 * TemplateRenderer is responsible for rendering a Google Docs
 * template into plain text by replacing placeholder variables.
 *
 * Responsibilities:
 * - Create a temporary copy of the Google Docs template
 * - Replace all placeholders using the `$Variable$` syntax
 * - Extract the final plain text content
 * - Ensure temporary resources are always cleaned up
 *
 * This class is intentionally stateless.
 */
class TemplateRenderer {

  /**
   * Renders a Google Docs template into plain text.
   *
   * The template may contain placeholders using the following syntax:
   *   - $Topic1$
   *   - $ Topic1 $
   *
   * Placeholders are replaced using the provided variables map.
   *
   * Implementation details:
   * - A temporary copy of the template is created to avoid mutating the original
   * - The copy is always deleted (trashed) in a finally block
   *
   * @param {string} templateId
   *   Google Docs file ID of the template.
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

      Object.keys(varsMap).forEach((key) => {
        const value = varsMap[key] == null ? '' : String(varsMap[key]);
        const escapedKey = Utils.escapeRegex(key);

        // Matches "$Topic1$" and "$ Topic1 $"
        const pattern = `\\$\\s*${escapedKey}\\s*\\$`;
        body.replaceText(pattern, value);
      });

      doc.saveAndClose();

      // Reopen to ensure we read the final committed content
      return DocumentApp
        .openById(copy.getId())
        .getBody()
        .getText();

    } finally {
      // Always clean up the temporary document to avoid Drive pollution
      copy.setTrashed(true);
    }
  }
}
