using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using MailMergeKit.WordAddin.Models;

namespace MailMergeKit.WordAddin.Services
{
    /// <summary>
    /// Controls the merge process: reads Word data source, processes records sequentially
    /// </summary>
    public class MergeController
    {
        private readonly OutlookMailer _mailer;
        private Word.Document _document;

        public MergeController()
        {
            _mailer = new OutlookMailer();
        }

        /// <summary>
        /// Starts the merge process for the active Word document
        /// </summary>
        /// <param name="doc">Active Word document with mail merge data source</param>
        /// <param name="subjectTemplate">Subject template with merge fields (e.g., "Hello «Name»")</param>
        /// <param name="recipientFieldName">Name of the field containing email addresses</param>
        /// <returns>Summary of merge operation</returns>
        public MergeResult StartMerge(Word.Document doc, string subjectTemplate, string recipientFieldName = "Email")
        {
            _document = doc;
            var result = new MergeResult();

            try
            {
                // Validate Word document has mail merge setup
                if (doc.MailMerge == null || doc.MailMerge.DataSource == null)
                {
                    result.ErrorMessage = "No mail merge data source found. Please set up mail merge first.";
                    return result;
                }

                // Check if Outlook is available
                if (!_mailer.IsOutlookAvailable())
                {
                    result.ErrorMessage = "Outlook is not running. Please start Outlook and try again.";
                    return result;
                }

                var dataSource = doc.MailMerge.DataSource;
                result.TotalRecords = dataSource.RecordCount;

                LogInfo($"Starting merge for {result.TotalRecords} recipients...");

                // Process each record sequentially (COM is single-threaded!)
                for (int i = 1; i <= dataSource.RecordCount; i++)
                {
                    dataSource.ActiveRecord = (Word.WdMailMergeActiveRecord)i;

                    try
                    {
                        // Build recipient data
                        var recipient = BuildRecipientData(dataSource, subjectTemplate, recipientFieldName, i);

                        if (recipient.IsValid())
                        {
                            // Create draft email
                            bool success = _mailer.CreateDraftEmail(recipient);

                            if (success)
                                result.SuccessCount++;
                            else
                                result.FailureCount++;
                        }
                        else
                        {
                            LogError($"Invalid data for row {i}: {recipient}");
                            result.FailureCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogError($"Error processing row {i}: {ex.Message}");
                        result.FailureCount++;
                    }

                    // Update progress
                    result.ProcessedRecords = i;
                }

                result.IsSuccess = result.SuccessCount > 0;
                LogInfo($"Merge complete: {result.SuccessCount} success, {result.FailureCount} failed");
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Merge failed: {ex.Message}";
                LogError(result.ErrorMessage);
            }

            return result;
        }

        /// <summary>
        /// Builds RecipientData from the current data source record
        /// </summary>
        private RecipientData BuildRecipientData(Word.MailMergeDataSource dataSource, 
            string subjectTemplate, string recipientFieldName, int rowNumber)
        {
            var recipient = new RecipientData
            {
                RowNumber = rowNumber,
                Email = GetFieldValue(dataSource, recipientFieldName),
                Subject = MergeFieldsInText(subjectTemplate, dataSource),
                BodyHtml = GetMergedBody(),
                CC = GetFieldValue(dataSource, "CC"),
                BCC = GetFieldValue(dataSource, "BCC"),
                Attachment = GetFieldValue(dataSource, "Attachment")
            };

            return recipient;
        }

        /// <summary>
        /// Gets the merged body HTML from Word document
        /// </summary>
        private string GetMergedBody()
        {
            try
            {
                // Get the HTML body of the current merged document
                // Word automatically merges fields when we access the content
                var range = _document.Content;
                return range.Text; // For v0.0.1, use text (HTML support in v0.1.0)
            }
            catch (Exception ex)
            {
                LogError($"Failed to get merged body: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Replaces merge fields in text (e.g., «Name» → John)
        /// </summary>
        private string MergeFieldsInText(string template, Word.MailMergeDataSource dataSource)
        {
            if (string.IsNullOrWhiteSpace(template))
                return template;

            // Match Word merge field format: «FieldName»
            var regex = new Regex(@"«([^»]+)»");
            
            return regex.Replace(template, match =>
            {
                var fieldName = match.Groups[1].Value;
                return GetFieldValue(dataSource, fieldName);
            });
        }

        /// <summary>
        /// Gets a field value from the data source
        /// </summary>
        private string GetFieldValue(Word.MailMergeDataSource dataSource, string fieldName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(fieldName))
                    return string.Empty;

                // Check if field exists
                foreach (Word.MailMergeDataField field in dataSource.DataFields)
                {
                    if (field.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        return field.Value?.ToString() ?? string.Empty;
                    }
                }

                return string.Empty;
            }
            catch (Exception ex)
            {
                LogError($"Failed to get field '{fieldName}': {ex.Message}");
                return string.Empty;
            }
        }

        #region Logging

        private void LogInfo(string message)
        {
            Console.WriteLine($"[INFO] {DateTime.Now:HH:mm:ss} | {message}");
        }

        private void LogError(string message)
        {
            Console.WriteLine($"[ERROR] {DateTime.Now:HH:mm:ss} | {message}");
        }

        #endregion
    }

    /// <summary>
    /// Result of a merge operation
    /// </summary>
    public class MergeResult
    {
        public bool IsSuccess { get; set; }
        public int TotalRecords { get; set; }
        public int ProcessedRecords { get; set; }
        public int SuccessCount { get; set; }
        public int FailureCount { get; set; }
        public string ErrorMessage { get; set; }

        public override string ToString()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
                return ErrorMessage;

            return $"Processed {ProcessedRecords}/{TotalRecords}: {SuccessCount} success, {FailureCount} failed";
        }
    }
}
