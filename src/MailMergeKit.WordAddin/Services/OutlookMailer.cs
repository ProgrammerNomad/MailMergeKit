using System;
using System.IO;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using MailMergeKit.WordAddin.Models;

namespace MailMergeKit.WordAddin.Services
{
    /// <summary>
    /// Handles Outlook COM operations for creating draft emails
    /// CRITICAL: COM is single-threaded (STA) - NO parallel processing!
    /// </summary>
    public class OutlookMailer : IDisposable
    {
        private static Outlook.Application _outlookApp;
        private static readonly object _lockObj = new object();
        private bool _disposed = false;

        /// <summary>
        /// Gets or creates the Outlook Application instance (singleton pattern for COM safety)
        /// </summary>
        private Outlook.Application GetOutlookApplication()
        {
            if (_outlookApp == null)
            {
                lock (_lockObj)
                {
                    if (_outlookApp == null)
                    {
                        try
                        {
                            // Try to get running instance first
                            _outlookApp = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
                        }
                        catch
                        {
                            // If not running, create new instance
                            _outlookApp = new Outlook.Application();
                        }
                    }
                }
            }
            return _outlookApp;
        }

        /// <summary>
        /// Creates a single draft email in Outlook Drafts folder
        /// </summary>
        /// <param name="recipient">Recipient data with all merge fields resolved</param>
        /// <returns>True if successful, false otherwise</returns>
        public bool CreateDraftEmail(RecipientData recipient)
        {
            if (recipient == null || !recipient.IsValid())
            {
                LogError($"Invalid recipient data: {recipient}");
                return false;
            }

            Outlook.MailItem mail = null;

            try
            {
                var outlookApp = GetOutlookApplication();
                
                // Create new mail item
                mail = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                // Set basic properties
                mail.To = recipient.Email;
                mail.Subject = recipient.Subject;
                mail.HTMLBody = recipient.BodyHtml;

                // Set CC/BCC if present
                if (!string.IsNullOrWhiteSpace(recipient.CC))
                    mail.CC = recipient.CC;

                if (!string.IsNullOrWhiteSpace(recipient.BCC))
                    mail.BCC = recipient.BCC;

                // Add attachments if specified
                if (!string.IsNullOrWhiteSpace(recipient.Attachment))
                {
                    AddAttachments(mail, recipient.Attachment, recipient.RowNumber);
                }

                // Save to Drafts folder (Outlook's default)
                // DO NOT use mail.Move() - it causes COM issues
                mail.Save();

                LogSuccess($"Draft created for {recipient.Email}");
                return true;
            }
            catch (Exception ex)
            {
                LogError($"Failed to create draft for {recipient.Email}: {ex.Message}");
                return false;
            }
            finally
            {
                // CRITICAL: Release COM object to prevent memory leaks
                if (mail != null)
                {
                    Marshal.ReleaseComObject(mail);
                    mail = null;
                }
            }
        }

        /// <summary>
        /// Adds attachments to the mail item (supports multiple files separated by semicolon)
        /// </summary>
        private void AddAttachments(Outlook.MailItem mail, string attachmentPaths, int rowNumber)
        {
            if (string.IsNullOrWhiteSpace(attachmentPaths))
                return;

            // Split by semicolon to support multiple attachments
            var files = attachmentPaths.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var file in files)
            {
                var filePath = file.Trim();

                // Resolve path (support relative paths)
                if (!Path.IsPathRooted(filePath))
                {
                    // Try relative to current directory
                    var fullPath = Path.Combine(Environment.CurrentDirectory, filePath);
                    if (File.Exists(fullPath))
                        filePath = fullPath;
                }

                if (File.Exists(filePath))
                {
                    try
                    {
                        mail.Attachments.Add(filePath, Outlook.OlAttachmentType.olByValue, 1, Path.GetFileName(filePath));
                        LogSuccess($"  Attached: {Path.GetFileName(filePath)}");
                    }
                    catch (Exception ex)
                    {
                        LogError($"  Failed to attach {filePath}: {ex.Message}");
                    }
                }
                else
                {
                    LogError($"  Attachment not found (Row {rowNumber}): {filePath}");
                }
            }
        }

        /// <summary>
        /// Checks if Outlook is running or can be started
        /// </summary>
        public bool IsOutlookAvailable()
        {
            try
            {
                var app = GetOutlookApplication();
                return app != null;
            }
            catch
            {
                return false;
            }
        }

        #region Logging (Simple Console for v0.0.1)

        private void LogSuccess(string message)
        {
            Console.WriteLine($"[SUCCESS] {DateTime.Now:HH:mm:ss} | {message}");
        }

        private void LogError(string message)
        {
            Console.WriteLine($"[ERROR] {DateTime.Now:HH:mm:ss} | {message}");
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Managed cleanup (none in this case)
                }

                // COM cleanup
                if (_outlookApp != null)
                {
                    // Note: We intentionally keep Outlook running
                    // as it's typically a user-managed application
                    // Just release our reference
                    Marshal.ReleaseComObject(_outlookApp);
                    _outlookApp = null;
                }

                _disposed = true;
            }
        }

        ~OutlookMailer()
        {
            Dispose(false);
        }

        #endregion
    }
}
