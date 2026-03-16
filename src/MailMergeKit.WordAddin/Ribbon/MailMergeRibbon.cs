using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using MailMergeKit.WordAddin.Services;
using MailMergeKit.WordAddin.UI;

namespace MailMergeKit.WordAddin.Ribbon
{
    [ComVisible(true)]
    public partial class MailMergeRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public MailMergeRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MailMergeKit.WordAddin.Ribbon.MailMergeRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        /// <summary>
        /// Ribbon load callback
        /// </summary>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        /// <summary>
        /// Button click handler for "Send via MailMergeKit"
        /// </summary>
        public void OnSendViaMailMergeKit(Office.IRibbonControl control)
        {
            try
            {
                var wordApp = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                var doc = wordApp.ActiveDocument;

                // Validate document has mail merge setup
                if (doc.MailMerge == null || doc.MailMerge.DataSource == null)
                {
                    MessageBox.Show(
                        "No mail merge data source found.\n\n" +
                        "Please set up mail merge first:\n" +
                        "1. Go to Mailings tab\n" +
                        "2. Click 'Select Recipients'\n" +
                        "3. Choose your data source (Excel, CSV, etc.)",
                        "MailMergeKit",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                // Show settings dialog
                using (var settingsForm = new SettingsForm(doc))
                {
                    if (settingsForm.ShowDialog() == DialogResult.OK)
                    {
                        // Get settings from form
                        var subjectTemplate = settingsForm.SubjectTemplate;
                        var recipientField = settingsForm.RecipientFieldName;

                        // Start merge
                        var controller = new MergeController();
                        var result = controller.StartMerge(doc, subjectTemplate, recipientField);

                        // Show result
                        if (result.IsSuccess)
                        {
                            MessageBox.Show(
                                $"Merge completed successfully!\n\n" +
                                $"Total: {result.TotalRecords}\n" +
                                $"Success: {result.SuccessCount}\n" +
                                $"Failed: {result.FailureCount}\n\n" +
                                $"Check your Outlook Drafts folder to review emails before sending.",
                                "MailMergeKit - Success",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show(
                                $"Merge completed with errors:\n\n{result}",
                                "MailMergeKit - Warning",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"An error occurred:\n\n{ex.Message}",
                    "MailMergeKit - Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
