using System;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace MailMergeKit.WordAddin
{
    /// <summary>
    /// VSTO Add-in entry point for MailMergeKit
    /// </summary>
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Add-in initialization
            LogStartup();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Cleanup on shutdown
            LogShutdown();
        }

        private void LogStartup()
        {
            try
            {
                Console.WriteLine($"[INFO] {DateTime.Now:HH:mm:ss} | MailMergeKit v0.0.1 loaded");
            }
            catch
            {
                // Suppress startup logging errors
            }
        }

        private void LogShutdown()
        {
            try
            {
                Console.WriteLine($"[INFO] {DateTime.Now:HH:mm:ss} | MailMergeKit shutting down");
            }
            catch
            {
                // Suppress shutdown logging errors
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
