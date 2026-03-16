using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace MailMergeKit.WordAddin.UI
{
    /// <summary>
    /// Settings dialog for configuring mail merge parameters
    /// </summary>
    public partial class SettingsForm : Form
    {
        private Word.Document _document;

        public string SubjectTemplate { get; private set; }
        public string RecipientFieldName { get; private set; }

        public SettingsForm(Word.Document document)
        {
            InitializeComponent();
            _document = document;
            LoadDataSourceFields();
            SetDefaults();
        }

        /// <summary>
        /// Loads available fields from the mail merge data source
        /// </summary>
        private void LoadDataSourceFields()
        {
            try
            {
                var dataSource = _document.MailMerge.DataSource;
                var fields = new List<string>();

                foreach (Word.MailMergeDataField field in dataSource.DataFields)
                {
                    fields.Add(field.Name);
                }

                cboRecipientField.Items.Clear();
                cboRecipientField.Items.AddRange(fields.ToArray());

                lblRecordCount.Text = $"Total records: {dataSource.RecordCount}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to load data source fields:\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Sets default values
        /// </summary>
        private void SetDefaults()
        {
            // Try to find "Email" field
            for (int i = 0; i < cboRecipientField.Items.Count; i++)
            {
                var fieldName = cboRecipientField.Items[i].ToString();
                if (fieldName.Equals("Email", StringComparison.OrdinalIgnoreCase) ||
                    fieldName.Equals("EmailAddress", StringComparison.OrdinalIgnoreCase) ||
                    fieldName.Equals("E-mail", StringComparison.OrdinalIgnoreCase))
                {
                    cboRecipientField.SelectedIndex = i;
                    break;
                }
            }

            // Default subject template
            txtSubject.Text = "Mail from MailMergeKit";
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            // Validate inputs
            if (cboRecipientField.SelectedIndex == -1)
            {
                MessageBox.Show(
                    "Please select the field containing recipient email addresses.",
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                cboRecipientField.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtSubject.Text))
            {
                MessageBox.Show(
                    "Please enter a subject line template.",
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                txtSubject.Focus();
                return;
            }

            // Save settings
            RecipientFieldName = cboRecipientField.SelectedItem.ToString();
            SubjectTemplate = txtSubject.Text;

            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void lnkMergeFieldHelp_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show(
                "Use Word merge field syntax in your subject:\n\n" +
                "Example: Hello «FirstName», your domain «Domain» expires on «ExpiryDate»\n\n" +
                "Available fields are listed in the dropdown above.\n" +
                "To insert a field, type «FieldName» where FieldName matches your data source column.",
                "Merge Field Help",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }
}
