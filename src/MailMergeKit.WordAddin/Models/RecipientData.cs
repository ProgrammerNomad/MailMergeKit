using System;

namespace MailMergeKit.WordAddin.Models
{
    /// <summary>
    /// Represents a single email recipient and their personalized data
    /// </summary>
    public class RecipientData
    {
        /// <summary>
        /// Recipient email address (required)
        /// </summary>
        public string Email { get; set; }

        /// <summary>
        /// Personalized subject line (merge fields already replaced)
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Email body HTML (merge fields already replaced)
        /// </summary>
        public string BodyHtml { get; set; }

        /// <summary>
        /// CC recipients (optional, semicolon-separated)
        /// </summary>
        public string CC { get; set; }

        /// <summary>
        /// BCC recipients (optional, semicolon-separated)
        /// </summary>
        public string BCC { get; set; }

        /// <summary>
        /// Attachment file paths (optional, semicolon-separated for multiple files)
        /// Example: "invoice.pdf" or "invoice.pdf;receipt.pdf"
        /// </summary>
        public string Attachment { get; set; }

        /// <summary>
        /// Row number in data source (for logging and error tracking)
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// Validates that required fields are present
        /// </summary>
        public bool IsValid()
        {
            return !string.IsNullOrWhiteSpace(Email) &&
                   !string.IsNullOrWhiteSpace(Subject) &&
                   !string.IsNullOrWhiteSpace(BodyHtml);
        }

        /// <summary>
        /// Returns a display-friendly representation for logging
        /// </summary>
        public override string ToString()
        {
            return $"Row {RowNumber}: {Email}";
        }
    }
}
